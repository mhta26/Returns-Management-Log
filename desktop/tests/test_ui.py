import os,time
from datetime import timedelta
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import pytest
from PySide6.QtWidgets import QApplication
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtGui import QPageSize,QPageLayout
from pypdf import PdfReader
from returns_app import ui
from returns_app.domain import ReceiptInput,now
from returns_app.store import Store


def wait(app,predicate):
    deadline=time.monotonic()+12
    while not predicate():
        if time.monotonic()>deadline:raise AssertionError('UI operation timed out')
        app.processEvents();time.sleep(.01)


class FakeApi:
    url='http://test-server'
    lose_ack=False
    def __init__(self,store):self.store=store
    def call(self,method,path,data=None,**kwargs):
        if path=='/api/session':return {'settings':self.store.settings(),'server_time':now().isoformat()}
        if method=='POST' and path=='/api/receipts':
            result=self.store.create_receipt(ReceiptInput(**data))
            if self.lose_ack:self.lose_ack=False;raise RuntimeError('Reply lost after commit')
            return result
        if path.endswith('/pdf'):return self.store.receipt(int(path.split('/')[3]))[1]
        if path.endswith('/print-events'):
            self.store.print_event(int(path.split('/')[3]),**data);return {'ok':True}
        return {'items':[],'total':0}


@pytest.fixture(scope='module')
def app():
    app=QApplication.instance() or QApplication([])
    app.setApplicationName('ReturnsManagerTest')
    return app


@pytest.fixture
def window(app,tmp_path,monkeypatch):
    folder=tmp_path/'client';folder.mkdir();monkeypatch.setattr(ui,'app_dir',lambda:folder)
    store=Store(tmp_path/'server.sqlite');api=FakeApi(store)
    w=ui.MainWindow(api,{'username':'مطيع','display_name':'مطيع','role':'admin'});w.error=lambda _:None;w.after_pdf=lambda *a:None
    wait(app,lambda:w.settings is not None)
    w.customer=store.search('customers')['items'][0];product=store.search('products')['items'][0]
    w.lines=[{'product':product,'cartons':1,'boxes':0,'pieces':0,'expiry':(now().date()+timedelta(days=60)).isoformat(),'reason':'','notes':''}];w.update_people();w.render_lines()
    yield w,api,store
    wait(app,lambda:not w.tasks.callbacks)
    w.close()


def test_lost_ack_keeps_draft_and_retry_does_not_duplicate(app,window):
    w,api,store=window;api.lose_ack=True;w.submit(False);wait(app,lambda:w.save_btn.isEnabled())
    assert w.pending and not w.entry.isEnabled() and store.list_receipts()['total']==1
    pending_id=w.pending['request_id'];w.save_draft()
    import json
    assert json.loads(w.draft_path.read_text(encoding='utf-8'))['pending']['request_id']==pending_id
    w.submit(False);wait(app,lambda:w.pending is None and w.save_btn.isEnabled())
    assert store.list_receipts()['total']==1 and not w.lines


def test_qt_print_backend_emits_requested_two_copies(app,window,tmp_path):
    w,api,store=window
    assert 'IBM Plex Sans Arabic' in app.font().family()
    receipt=store.create_receipt(ReceiptInput(**w.payload()));pdf=store.receipt(receipt['serial'])[1]
    printer=QPrinter(QPrinter.PrinterMode.HighResolution);printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat);printer.setOutputFileName(str(tmp_path/'copies.pdf'));printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4));printer.setPageOrientation(QPageLayout.Orientation.Landscape);printer.setCopyCount(2)
    w.print_content(receipt['serial'],pdf,printer);wait(app,lambda:not w.tasks.callbacks)
    assert len(PdfReader(tmp_path/'copies.pdf').pages)==2
    with store.connect() as c:assert c.execute('SELECT copies,success FROM print_events').fetchone()['copies']==2


def test_search_does_not_select_stale_results_after_typing(app,window):
    w,api,store=window
    dialog=ui.SearchDialog(api,'customers',w)
    wait(app,lambda:not dialog.tasks.callbacks)
    dialog.rows=[store.search('customers')['items'][0]]
    ui.fill(dialog.grid,[[dialog.rows[0]['name'],dialog.rows[0]['code']]])
    dialog.grid.selectRow(0)
    dialog.query.setText('بحث مختلف قبل اكتمال الطلب')
    dialog.choose()
    assert dialog.result()!=dialog.DialogCode.Accepted and dialog.selected is None
    wait(app,lambda:not dialog.tasks.callbacks)
    dialog.close()
