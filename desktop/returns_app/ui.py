from __future__ import annotations
import json
import os
import re
import sys
import uuid
from pathlib import Path
from datetime import date
from urllib.parse import urlencode
from PySide6.QtCore import Qt,QDate,QTimer,QStandardPaths,QUrl,QByteArray,QBuffer,QIODevice,QSize,QRectF
from PySide6.QtGui import QColor,QPainter,QDesktopServices,QFont,QPageSize,QPageLayout,QPixmap,QIcon
from PySide6.QtWidgets import (QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QGridLayout,QFormLayout,QLabel,QPushButton,QLineEdit,QDateEdit,QSpinBox,QComboBox,QTableWidget,QTableWidgetItem,QHeaderView,QAbstractItemView,QTabWidget,QListWidget,QStackedWidget,QDialog,QDialogButtonBox,QMessageBox,QFileDialog,QInputDialog,QTextEdit,QSplitter,QFrame,QCheckBox)
from PySide6.QtPrintSupport import QPrinter,QPrintDialog
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from .client import Api,Tasks
from .domain import status_for,now
from .pdf import PALETTE,ASSETS

STYLE='''
QWidget{font-family:"Segoe UI","DejaVu Sans";font-size:13px;color:#152c42;}
QMainWindow,QDialog{background:#f3f6fa;}
QLineEdit,QComboBox,QDateEdit,QSpinBox,QTextEdit{background:white;border:1px solid #cdd7e3;border-radius:7px;padding:7px;min-height:23px;}
QLineEdit:focus,QComboBox:focus{border:1px solid #087e8b;}
QPushButton{background:#143e68;color:white;border:0;border-radius:7px;padding:9px 15px;min-height:23px;}
QPushButton:hover{background:#1b568d;}QPushButton:disabled{background:#a7b8c9;}
QPushButton[secondary="true"]{background:#e2ebf3;color:#143e68;}
QListWidget{background:#102f50;color:white;border:0;border-radius:10px;padding:8px;}
QListWidget::item{padding:15px 10px;border-radius:6px;margin-bottom:5px;color:white;}
QListWidget::item:selected{background:#087e8b;}
QTableWidget{background:white;border:1px solid #d7e0e9;gridline-color:#e6edf3;selection-background-color:#d7edf1;selection-color:#153952;}
QHeaderView::section{background:#143e68;color:white;padding:9px;border:0;}
QLabel#title{font-size:25px;font-weight:700;color:#143e68;}QLabel#subtitle{color:#536c83;}
QLabel#badge{background:#e1f3ef;color:#12684e;border-radius:8px;padding:7px;}
QFrame#card{background:white;border:1px solid #dce5ed;border-radius:10px;}
QTabWidget::pane{border:1px solid #d8e1ea;background:white;}QTabBar::tab{padding:12px;background:#e1eaf3;}QTabBar::tab:selected{background:#143e68;color:white;}
'''


def button(text,slot=None,secondary=False):
    b=QPushButton(text);b.setProperty('secondary',secondary)
    if slot:b.clicked.connect(slot)
    return b


def table(headers):
    t=QTableWidget(0,len(headers));t.setHorizontalHeaderLabels(headers)
    t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    t.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    t.verticalHeader().setVisible(False);t.horizontalHeader().setStretchLastSection(True)
    return t


def fill(t,rows):
    t.setRowCount(len(rows))
    for r,row in enumerate(rows):
        for c,value in enumerate(row):t.setItem(r,c,QTableWidgetItem(str(value if value is not None else '')))
    t.resizeColumnsToContents()
    for c in range(t.columnCount()):
        if t.columnWidth(c)>360:t.setColumnWidth(c,360)


def app_dir():
    p=Path(QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppLocalDataLocation));p.mkdir(parents=True,exist_ok=True);return p


def atomic_json(path,value):
    temp=path.with_suffix('.tmp');temp.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8');os.replace(temp,path)


class LoginDialog(QDialog):
    def __init__(self,parent=None):
        super().__init__(parent);self.setWindowTitle('تسجيل الدخول • إدارة المرتجعات');self.resize(490,360)
        v=QVBoxLayout(self);title=QLabel('مرحبًا بك');title.setObjectName('title');v.addWidget(title)
        v.addWidget(QLabel('بيانات موحّدة لكل الأجهزة • حساب مستقل لكل أمين صندوق'))
        self.config_path=app_dir()/'connection.json'
        try:config=json.loads(self.config_path.read_text(encoding='utf-8'))
        except (OSError,ValueError):config={}
        f=QFormLayout();self.url=QLineEdit(config.get('url','http://127.0.0.1:8765'));self.url.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.user=QLineEdit(config.get('username','مطيع'));self.password=QLineEdit();self.password.setEchoMode(QLineEdit.EchoMode.Password)
        f.addRow('عنوان الخادم',self.url);f.addRow('اسم المستخدم',self.user);f.addRow('كلمة المرور',self.password);v.addLayout(f)
        self.message=QLabel('الحساب الأول: مطيع. كلمة المرور الأولية في ملف إعداد الخادم.');self.message.setWordWrap(True);v.addWidget(self.message)
        self.go=button('دخول',self.login);v.addWidget(self.go);self.tasks=Tasks(self);self.password.returnPressed.connect(self.login)
    def login(self):
        try:self.api=Api(self.url.text().strip())
        except ValueError as e:self.message.setText(str(e));return
        self.go.setEnabled(False);self.message.setText('جارٍ الاتصال…')
        username,password=self.user.text().strip(),self.password.text()
        self.tasks.submit(lambda:self.api.call('POST','/auth/login',{'username':username,'password':password}),self.success,self.failure)
    def success(self,result):
        self.api.token=result['token'];self.identity=result['identity'];atomic_json(self.config_path,{'url':self.api.url,'username':self.user.text().strip()});self.accept()
    def failure(self,error):self.go.setEnabled(True);self.message.setText(error)


class SearchDialog(QDialog):
    def __init__(self,api,kind,parent=None):
        super().__init__(parent);self.api=api;self.kind=kind;self.selected=None;self.generation=0;self.offset=0;self.rows=[]
        self.setWindowTitle({'customers':'اختيار العميل','products':'اختيار الصنف','drivers':'اختيار السائق'}[kind]);self.resize(850,530)
        v=QVBoxLayout(self);self.query=QLineEdit();self.query.setPlaceholderText('اكتب جزءًا من الاسم أو الكود…');v.addWidget(self.query)
        self.grid=table(['الاسم','الكود']);v.addWidget(self.grid);self.count=QLabel();v.addWidget(self.count)
        h=QHBoxLayout();h.addWidget(button('السابق',lambda:self.page(-1),True));h.addWidget(button('التالي',lambda:self.page(1),True));h.addStretch();h.addWidget(button('اختيار',self.choose));h.addWidget(button('إلغاء',self.reject,True));v.addLayout(h)
        self.tasks=Tasks(self);self.timer=QTimer(self);self.timer.setSingleShot(True);self.timer.setInterval(220)
        self.query.textChanged.connect(self.schedule);self.timer.timeout.connect(self.load);self.grid.doubleClicked.connect(self.choose);self.query.returnPressed.connect(self.choose);self.load()
    def schedule(self):self.offset=0;self.timer.start()
    def page(self,delta):self.offset=max(0,self.offset+delta*100);self.load()
    def load(self):
        self.generation+=1;generation=self.generation;path='/api/masters/'+self.kind+'?'+urlencode({'q':self.query.text(),'offset':self.offset})
        self.tasks.submit(lambda:self.api.call('GET',path),lambda r:self.loaded(r,generation),lambda e:self.count.setText(e))
    def loaded(self,result,generation):
        if generation!=self.generation:return
        self.rows=result['items'];fill(self.grid,[[r['name'],r['code']] for r in self.rows]);self.count.setText(f"{self.offset+1 if self.rows else 0}–{self.offset+len(self.rows)} من {result['total']}")
        if self.rows:self.grid.selectRow(0)
    def choose(self):
        row=self.grid.currentRow()
        if 0<=row<len(self.rows):self.selected=self.rows[row];self.accept()


class PdfDialog(QDialog):
    def __init__(self,content,title,parent=None):
        super().__init__(parent);self.setWindowTitle(title);self.resize(1100,750);v=QVBoxLayout(self)
        self.buffer=QBuffer(self);self.buffer.setData(QByteArray(content));self.buffer.open(QIODevice.OpenModeFlag.ReadOnly)
        self.doc=QPdfDocument(self);self.doc.load(self.buffer)
        self.view=QPdfView(self);self.view.setDocument(self.doc);self.view.setPageMode(QPdfView.PageMode.MultiPage);self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth);v.addWidget(self.view)
        v.addWidget(button('إغلاق',self.accept,True))


class MainWindow(QMainWindow):
    def __init__(self,api,identity):
        super().__init__();self.api=api;self.identity=identity;self.tasks=Tasks(self);self.settings=None;self.customer=None;self.driver=None;self.lines=[];self.pending=None;self.request_id=str(uuid.uuid4());self.restoring=False;self.draft_path=app_dir()/('draft-'+identity['username']+'.json');self.pending_printer=None
        self.setWindowTitle('إدارة المرتجعات • Shanifi');self.resize(1440,910);self.setMinimumSize(1050,700)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        root=QWidget();self.setCentralWidget(root);main=QHBoxLayout(root);main.setContentsMargins(16,16,16,16)
        sidebar=QVBoxLayout();logo=QLabel();logo.setPixmap(QPixmap(str(ASSETS/'company.jpeg')).scaledToWidth(205,Qt.TransformationMode.SmoothTransformation));sidebar.addWidget(logo)
        self.user_badge=QLabel('أمين الصندوق: '+identity['display_name']);self.user_badge.setObjectName('badge');self.user_badge.setWordWrap(True);sidebar.addWidget(self.user_badge)
        self.nav=QListWidget();self.nav.setFixedWidth(230);self.nav.addItems(['سند مرتجع جديد','الأرشيف والسندات','قاعدة العملاء','دليل الأصناف','قاعدة السائقين','الإعدادات والنسخ الاحتياطي','المستخدمون وأمناء الصندوق']);sidebar.addWidget(self.nav)
        sidebar.addWidget(QLabel('الخادم: '+api.url));sidebar.addWidget(button('تسجيل الخروج',self.logout,True));main.addLayout(sidebar)
        self.pages=QStackedWidget();main.addWidget(self.pages,1)
        self.build_entry();self.build_archive();self.master_pages={}
        for kind,title in [('customers','قاعدة العملاء'),('products','دليل الأصناف'),('drivers','قاعدة السائقين')]:self.build_masters(kind,title)
        self.build_settings();self.build_users();self.nav.currentRowChanged.connect(self.navigate);self.nav.setCurrentRow(0)
        self.statusBar().showMessage('جارٍ تحميل الإعدادات…');self.entry.setEnabled(False)
        self.autosave=QTimer(self);self.autosave.setSingleShot(True);self.autosave.setInterval(450);self.autosave.timeout.connect(self.save_draft)
        self.tasks.submit(lambda:self.api.call('GET','/api/session'),self.loaded,self.error)

    def error(self,message):self.statusBar().showMessage(message,15000);QMessageBox.warning(self,'تعذر إكمال العملية',message)
    def title(self,text,subtitle=''):
        w=QWidget();v=QVBoxLayout(w);v.setContentsMargins(6,6,6,6);l=QLabel(text);l.setObjectName('title');v.addWidget(l)
        if subtitle:s=QLabel(subtitle);s.setObjectName('subtitle');s.setWordWrap(True);v.addWidget(s)
        return w,v

    def build_entry(self):
        page,v=self.title('سند مرتجع جديد','اختر العميل والسائق، ثم أضف الأصناف. تُحفظ البيانات مركزيًا ويُنشأ PDF قبل تجهيز السند التالي.')
        self.entry=QWidget();layout=QVBoxLayout(self.entry);layout.setContentsMargins(0,0,0,0)
        top=QHBoxLayout();self.serial_label=QLabel('التسلسل التالي: —');self.serial_label.setObjectName('badge');top.addWidget(self.serial_label);self.clock=QLabel();top.addWidget(self.clock);top.addStretch();layout.addLayout(top)
        g=QGridLayout();self.customer_text=QLineEdit();self.customer_text.setReadOnly(True);self.customer_text.setPlaceholderText('اختر العميل بالاسم أو الكود');g.addWidget(QLabel('العميل'),0,0);g.addWidget(self.customer_text,0,1,1,3);g.addWidget(button('بحث العملاء',lambda:self.select_master('customers')),0,4)
        self.customer_info=QLabel('المندوب والقطاع يظهران بعد اختيار العميل');self.customer_info.setWordWrap(True);g.addWidget(self.customer_info,1,1,1,4)
        self.driver_text=QLineEdit();self.driver_text.setReadOnly(True);g.addWidget(QLabel('السائق'),2,0);g.addWidget(self.driver_text,2,1);g.addWidget(button('اختيار السائق',lambda:self.select_master('drivers'),True),2,2)
        self.return_number=QLineEdit();self.return_number.setPlaceholderText('رقم الإرجاع');g.addWidget(QLabel('رقم الإرجاع'),2,3);g.addWidget(self.return_number,2,4)
        self.invoice_enabled=QCheckBox('يوجد تاريخ فاتورة');self.invoice_date=QDateEdit(QDate.currentDate());self.invoice_date.setCalendarPopup(True);self.invoice_date.setDisplayFormat('yyyy-MM-dd');self.invoice_date.setEnabled(False);self.invoice_enabled.toggled.connect(self.invoice_date.setEnabled);g.addWidget(self.invoice_enabled,3,0);g.addWidget(self.invoice_date,3,1)
        self.notes=QLineEdit();self.notes.setPlaceholderText('ملاحظة عامة للسند (اختياري)');g.addWidget(self.notes,3,2,1,3);layout.addLayout(g)
        h=QHBoxLayout();h.addWidget(button('＋ إضافة صنف',self.add_line));h.addWidget(button('تعديل الصنف',self.edit_line,True));h.addWidget(button('حذف الصنف',self.remove_line,True));h.addStretch();layout.addLayout(h)
        self.line_grid=table(['م','كود الصنف','اسم الصنف','كرتون','علبة','حبة','تاريخ الانتهاء','المخزن / الحالة','سبب الإرجاع','ملاحظات']);self.line_grid.doubleClicked.connect(self.edit_line);layout.addWidget(self.line_grid,1)
        self.totals=QLabel('لا توجد أصناف بعد');layout.addWidget(self.totals);v.addWidget(self.entry,1)
        actions=QHBoxLayout();self.print_btn=button('طباعة وحفظ PDF',lambda:self.submit(True));self.save_btn=button('حفظ السند وPDF',lambda:self.submit(False));self.clear_btn=button('سند جديد / تفريغ',self.clear_confirm,True)
        actions.addWidget(self.print_btn);actions.addWidget(self.save_btn);actions.addStretch();actions.addWidget(self.clear_btn);v.addLayout(actions);self.pages.addWidget(page)
        self.return_number.textChanged.connect(self.changed);self.invoice_enabled.toggled.connect(self.changed);self.invoice_date.dateChanged.connect(self.changed);self.notes.textChanged.connect(self.changed)

    def loaded(self,result):
        self.settings=result['settings'];self.clock.setText('تاريخ السند: '+result['server_time'][:10]);self.serial_label.setText('التسلسل التالي: '+str(self.settings['next_serial']));self.entry.setEnabled(True);self.restore_draft();self.render_lines();self.populate_settings();self.statusBar().showMessage('متصل • البيانات مشتركة • '+self.identity['display_name'])

    def navigate(self,index):
        if index<0:return
        self.pages.setCurrentIndex(index)
        if index==1:self.load_archive()
        elif index in [2,3,4]:self.load_masters(['customers','products','drivers'][index-2])
        elif index==6:self.load_users()

    def select_master(self,kind):
        dialog=SearchDialog(self.api,kind,self)
        if dialog.exec()==QDialog.DialogCode.Accepted:
            if kind=='customers':self.customer=dialog.selected
            else:self.driver=dialog.selected
            self.update_people();self.changed()
    def update_people(self):
        self.customer_text.setText((self.customer['name']+' | '+self.customer['code']) if self.customer else '')
        c=self.customer or {};self.customer_info.setText('المندوب: '+c.get('rep_code','')+' — '+c.get('rep_name','')+'    |    القطاع: '+c.get('sector',''))
        self.driver_text.setText((self.driver['name']+' | '+self.driver['code']) if self.driver else '')

    def line_dialog(self,value=None):
        if not self.settings:return None
        d=QDialog(self);d.setWindowTitle('بيانات الصنف');d.resize(580,520);v=QVBoxLayout(d);f=QFormLayout();product=(value or {}).get('product')
        selection=QLineEdit();selection.setReadOnly(True)
        if product:selection.setText(product['name']+' | '+product['code'])
        def choose():
            nonlocal product
            picker=SearchDialog(self.api,'products',d)
            if picker.exec()==QDialog.DialogCode.Accepted:product=picker.selected;selection.setText(product['name']+' | '+product['code'])
        row=QHBoxLayout();row.addWidget(selection,1);row.addWidget(button('بحث',choose));f.addRow('الصنف',row)
        quantities=[]
        for key,label in [('cartons','كرتون'),('boxes','علبة'),('pieces','حبة')]:
            spin=QSpinBox();spin.setRange(0,1000000000);spin.setValue((value or {}).get(key,0));f.addRow(label,spin);quantities.append(spin)
        expiry=QDateEdit();expiry.setDisplayFormat('yyyy-MM-dd');expiry.setCalendarPopup(True);expiry.setDate(QDate.fromString((value or {}).get('expiry',date.today().isoformat()),'yyyy-MM-dd'));f.addRow('تاريخ الانتهاء',expiry)
        reason=QComboBox();reason.addItems(['']+self.settings['reasons']);reason.setCurrentText((value or {}).get('reason',''));f.addRow('سبب الإرجاع',reason)
        notes=QTextEdit();notes.setMaximumHeight(90);notes.setPlainText((value or {}).get('notes',''));f.addRow('ملاحظات',notes);v.addLayout(f)
        info=QLabel();info.setObjectName('badge');v.addWidget(info)
        def classification():
            label,index,remaining=status_for(expiry.date().toPython(),now().date(),self.settings);info.setText(f'{label} • المتبقي: {remaining} يومًا');info.setStyleSheet('background:'+PALETTE[index]+';padding:9px;border-radius:6px')
        expiry.dateChanged.connect(classification);classification()
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel)
        def accept():
            if not product:QMessageBox.warning(d,'الصنف','اختر صنفًا من قاعدة البيانات.');return
            if sum(q.value() for q in quantities)<=0:QMessageBox.warning(d,'الكمية','أدخل كمية أكبر من صفر.');return
            d.accept()
        buttons.accepted.connect(accept);buttons.rejected.connect(d.reject);v.addWidget(buttons)
        if d.exec()!=QDialog.DialogCode.Accepted:return None
        return {'product':product,'cartons':quantities[0].value(),'boxes':quantities[1].value(),'pieces':quantities[2].value(),'expiry':expiry.date().toString('yyyy-MM-dd'),'reason':reason.currentText(),'notes':notes.toPlainText()}

    def add_line(self):
        if len(self.lines)>=220:self.error('الحد الأقصى 220 صنفًا للسند.');return
        value=self.line_dialog()
        if value:self.lines.append(value);self.render_lines();self.changed()
    def edit_line(self):
        row=self.line_grid.currentRow()
        if 0<=row<len(self.lines):
            value=self.line_dialog(self.lines[row])
            if value:self.lines[row]=value;self.render_lines();self.changed()
    def remove_line(self):
        row=self.line_grid.currentRow()
        if 0<=row<len(self.lines):self.lines.pop(row);self.render_lines();self.changed()
    def render_lines(self):
        if not self.settings:return
        rows=[];indices=[]
        for i,line in enumerate(self.lines,1):
            status,index,_=status_for(date.fromisoformat(line['expiry']),now().date(),self.settings);indices.append(index)
            rows.append([i,line['product']['code'],line['product']['name'],line['cartons'],line['boxes'],line['pieces'],line['expiry'],status,line['reason'],line['notes']])
        fill(self.line_grid,rows)
        for row,index in enumerate(indices):self.line_grid.item(row,7).setBackground(QColor(PALETTE[index]))
        self.totals.setText(f"عدد الأصناف: {len(self.lines)}     كرتون: {sum(l['cartons'] for l in self.lines)}     علبة: {sum(l['boxes'] for l in self.lines)}     حبة: {sum(l['pieces'] for l in self.lines)}")

    def changed(self,*args):
        if self.restoring:return
        self.request_id=str(uuid.uuid4())
        if hasattr(self,'autosave'):self.autosave.start()
    def save_draft(self):
        draft={'customer':self.customer,'driver':self.driver,'lines':self.lines,'return_number':self.return_number.text(),'invoice_enabled':self.invoice_enabled.isChecked(),'invoice_date':self.invoice_date.date().toString('yyyy-MM-dd'),'notes':self.notes.text(),'request_id':self.request_id,'pending':self.pending,'server':self.api.url}
        try:atomic_json(self.draft_path,draft)
        except OSError as e:self.statusBar().showMessage('تعذر حفظ المسودة محليًا: '+str(e))
    def restore_draft(self):
        try:draft=json.loads(self.draft_path.read_text(encoding='utf-8'))
        except (OSError,ValueError):return
        if draft.get('server')!=self.api.url:return
        self.restoring=True
        self.customer=draft.get('customer');self.driver=draft.get('driver');self.lines=draft.get('lines',[]);self.return_number.setText(draft.get('return_number',''));self.invoice_enabled.setChecked(draft.get('invoice_enabled',False));self.invoice_date.setDate(QDate.fromString(draft.get('invoice_date',date.today().isoformat()),'yyyy-MM-dd'));self.notes.setText(draft.get('notes',''));self.request_id=draft.get('request_id',str(uuid.uuid4()));self.pending=draft.get('pending');self.update_people();self.restoring=False
        if self.pending:self.entry.setEnabled(False);self.clear_btn.setEnabled(False);self.statusBar().showMessage('توجد عملية لم يُؤكَّد حفظها. اضغط حفظ لإعادة المحاولة بالمعرّف نفسه.')

    def payload(self):
        if not self.customer:raise ValueError('اختر العميل أولًا.')
        if not self.lines:raise ValueError('أضف صنفًا واحدًا على الأقل.')
        return {'request_id':self.request_id,'operator':self.identity['display_name'],'customer_code':self.customer['code'],'driver_code':(self.driver or {}).get('code',''),'return_number':self.return_number.text(),'invoice_date':self.invoice_date.date().toString('yyyy-MM-dd') if self.invoice_enabled.isChecked() else None,'notes':self.notes.text(),'items':[{'product_code':l['product']['code'],**{k:l[k] for k in ['cartons','boxes','pieces','expiry','reason','notes']}} for l in self.lines]}
    def choose_printer(self):
        count,ok=QInputDialog.getInt(self,'عدد النسخ','كم نسخة تريد طباعتها؟',1,1,999)
        if not ok:return None
        printer=QPrinter(QPrinter.PrinterMode.HighResolution);printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4));printer.setPageOrientation(QPageLayout.Orientation.Landscape);printer.setCopyCount(count);printer.setCollateCopies(True)
        dialog=QPrintDialog(printer,self);dialog.setWindowTitle('اختيار الطابعة')
        if dialog.exec()!=QDialog.DialogCode.Accepted:return None
        return printer
    def submit(self,print_now):
        if not self.settings:return
        try:payload=self.pending or self.payload()
        except ValueError as e:self.error(str(e));return
        printer=self.choose_printer() if print_now else None
        if print_now and printer is None:return
        self.pending=payload;self.pending_printer=printer;self.save_draft();self.entry.setEnabled(False);self.clear_btn.setEnabled(False);self.save_btn.setEnabled(False);self.print_btn.setEnabled(False);self.statusBar().showMessage('جارٍ حفظ السند وPDF في الخادم…')
        self.tasks.submit(lambda:self.api.call('POST','/api/receipts',payload),self.saved,self.save_failed)
    def save_failed(self,error):
        self.save_btn.setEnabled(True);self.print_btn.setEnabled(True);self.error(error+'\nاحتُفظ بالمدخلات ومعرّف العملية. أعد الضغط على حفظ لتأكيد النتيجة دون تكرار السند.')
    def saved(self,snapshot):
        printer=self.pending_printer;self.pending=None;self.reset_entry();self.save_btn.setEnabled(True);self.print_btn.setEnabled(True)
        self.settings['next_serial']=max(self.settings['next_serial'],snapshot['serial']+1);self.serial_label.setText('التسلسل التالي: '+str(self.settings['next_serial']));self.statusBar().showMessage('حُفظ السند #'+str(snapshot['serial'])+' مركزيًا؛ جارٍ جلب PDF…')
        self.tasks.submit(lambda:self.api.call('GET',f"/api/receipts/{snapshot['serial']}/pdf",binary=True),lambda content:self.after_pdf(snapshot,content,printer),lambda e:self.error('السند محفوظ في الأرشيف، لكن تعذر تنزيل PDF الآن. '+e))
    def reset_entry(self):
        self.restoring=True;self.customer=None;self.driver=None;self.lines=[];self.return_number.clear();self.invoice_enabled.setChecked(False);self.invoice_date.setDate(QDate.currentDate());self.notes.clear();self.request_id=str(uuid.uuid4());self.update_people();self.restoring=False;self.entry.setEnabled(True);self.clear_btn.setEnabled(True);self.render_lines();self.save_draft()
    def clear_confirm(self):
        if self.pending:return
        if QMessageBox.question(self,'تفريغ السند','تفريغ المدخلات الحالية غير المعتمدة؟')==QMessageBox.StandardButton.Yes:self.reset_entry()
    def after_pdf(self,snapshot,content,printer):
        try:
            folder=Path(QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation))/'ReturnsManager PDFs';folder.mkdir(parents=True,exist_ok=True)
            name=re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',snapshot['customer']['name']).strip('. ')[:100]+' - '+str(snapshot['serial'])+'.pdf';path=folder/name
            if path.exists() and path.read_bytes()!=content:path=folder/(str(snapshot['serial'])+'-'+uuid.uuid4().hex[:8]+'.pdf')
            if not path.exists():
                with path.open('xb') as f:f.write(content)
        except OSError as e:self.error('السند محفوظ مركزيًا؛ تعذر حفظ نسخة محلية: '+str(e))
        if printer:self.print_content(snapshot['serial'],content,printer)
        else:self.statusBar().showMessage('حُفظ السند وPDF. يمكنك فتحه من الأرشيف.',12000)
    def print_content(self,serial,content,printer):
        buffer=QBuffer();buffer.setData(QByteArray(content));buffer.open(QIODevice.OpenModeFlag.ReadOnly);doc=QPdfDocument();doc.load(buffer)
        painter=QPainter();error='';success=False;requested_copies=printer.copyCount()
        try:
            if doc.pageCount()<1:raise RuntimeError('تعذر قراءة PDF.')
            if not painter.begin(printer):raise RuntimeError('تعذر الاتصال بالطابعة.')
            # Render at 180dpi; the canonical archived PDF remains vector quality.
            manual_copies = 1 if printer.supportsMultipleCopies() else requested_copies
            if manual_copies>1:printer.setCopyCount(1)
            page_index=0
            for copy in range(manual_copies):
                for page in range(doc.pageCount()):
                    if page_index and not printer.newPage():raise RuntimeError('تعذر إنشاء صفحة في الطابعة.')
                    size=doc.pagePointSize(page);image=doc.render(page,QSize(round(size.width()*2.5),round(size.height()*2.5)))
                    if image.isNull():raise RuntimeError('تعذر تجهيز صفحة PDF للطباعة.')
                    area=QRectF(painter.viewport());ratio=min(area.width()/image.width(),area.height()/image.height());target=QRectF((area.width()-image.width()*ratio)/2,(area.height()-image.height()*ratio)/2,image.width()*ratio,image.height()*ratio);painter.drawImage(target,image)
                    page_index+=1
            success=True
        except Exception as e:error=str(e);self.error('السند محفوظ، لكن تعذرت الطباعة: '+error)
        finally:
            if painter.isActive():painter.end()
            doc.close();buffer.close()
        event={'operator':self.identity['display_name'],'printer':printer.printerName(),'copies':requested_copies,'success':success,'error':error}
        self.tasks.submit(lambda:self.api.call('POST',f'/api/receipts/{serial}/print-events',event),lambda _:self.statusBar().showMessage('أُرسل السند للطابعة.' if success else 'السند محفوظ؛ أعد الطباعة من الأرشيف.',12000),lambda e:self.error('حُفظ السند؛ تعذر تحديث حالة الطباعة مركزيًا. '+e))

    def build_archive(self):
        page,v=self.title('الأرشيف والسندات','نسخ ثابتة من كل سند • إعادة الطباعة لا تغيّر التسلسل أو التصنيف التاريخي')
        h=QHBoxLayout();self.archive_query=QLineEdit();self.archive_query.setPlaceholderText('بحث بالعميل أو السائق أو رقم الإرجاع أو التسلسل…');h.addWidget(self.archive_query,1);h.addWidget(button('بحث',lambda:self.load_archive(reset=True)));v.addLayout(h)
        dates=QHBoxLayout();self.filter_dates=QCheckBox('تصفية حسب التاريخ');self.archive_start=QDateEdit(QDate.currentDate().addMonths(-1));self.archive_end=QDateEdit(QDate.currentDate())
        for d in [self.archive_start,self.archive_end]:d.setCalendarPopup(True);d.setDisplayFormat('yyyy-MM-dd')
        dates.addWidget(self.filter_dates);dates.addWidget(QLabel('من'));dates.addWidget(self.archive_start);dates.addWidget(QLabel('إلى'));dates.addWidget(self.archive_end);dates.addStretch();v.addLayout(dates)
        self.archive_grid=table(['التسلسل','تاريخ السند','كود العميل','العميل','أمين الصندوق','عدد الأصناف','كرتون','علبة','حبة','رقم الإرجاع','تاريخ الفاتورة','السائق','المخازن','الطباعة']);v.addWidget(self.archive_grid,1)
        self.archive_rows=[];self.archive_offset=0;self.archive_count=QLabel();v.addWidget(self.archive_count)
        h=QHBoxLayout()
        for text,fn in [('معاينة PDF',lambda:self.archive_action('preview')),('إعادة الطباعة',lambda:self.archive_action('print')),('حفظ PDF',lambda:self.archive_action('download')),('تفاصيل السند',self.receipt_details),('تصدير CSV',self.export_csv),('إرفاق PDF قديم',self.attach_pdf)]:h.addWidget(button(text,fn,True))
        v.addLayout(h);h=QHBoxLayout();h.addWidget(button('السابق',lambda:self.archive_page(-1),True));h.addWidget(button('التالي',lambda:self.archive_page(1),True));h.addStretch();v.addLayout(h);self.archive_grid.doubleClicked.connect(lambda:self.archive_action('preview'));self.archive_query.returnPressed.connect(lambda:self.load_archive(reset=True));self.pages.addWidget(page)
    def archive_params(self):
        return {'q':self.archive_query.text(),'start':self.archive_start.date().toString('yyyy-MM-dd') if self.filter_dates.isChecked() else '', 'end':self.archive_end.date().toString('yyyy-MM-dd') if self.filter_dates.isChecked() else ''}
    def load_archive(self,reset=False):
        if reset:self.archive_offset=0
        path='/api/receipts?'+urlencode({**self.archive_params(),'offset':self.archive_offset})
        self.tasks.submit(lambda:self.api.call('GET',path),self.archive_loaded,self.error)
    def archive_loaded(self,result):
        self.archive_rows=result['items'];rows=[]
        for r in self.archive_rows:rows.append([r['serial'],r['created_at'][:16].replace('T',' '),r['customer']['code'],r['customer']['name'],r.get('operator',''),r.get('legacy_count') or len(r['items']),*[r['totals'][k] for k in ['cartons','boxes','pieces']],r.get('return_number',''),r.get('invoice_date',''),r['driver'].get('name',''),'، '.join(sorted(set(i['status'] for i in r['items']))) or r.get('legacy_warehouses',''),r['print_state']])
        fill(self.archive_grid,rows);self.archive_count.setText(f"{self.archive_offset+1 if rows else 0}–{self.archive_offset+len(rows)} من {result['total']} سند")
        if rows:self.archive_grid.selectRow(0)
    def archive_page(self,delta):self.archive_offset=max(0,self.archive_offset+100*delta);self.load_archive()
    def selected_receipt(self):
        row=self.archive_grid.currentRow()
        if 0<=row<len(self.archive_rows):return self.archive_rows[row]
        return None
    def archive_action(self,action):
        receipt=self.selected_receipt()
        if not receipt:return
        printer=self.choose_printer() if action=='print' else None
        if action=='print' and printer is None:return
        path=None
        if action=='download':
            path,_=QFileDialog.getSaveFileName(self,'حفظ PDF',f"سند - {receipt['serial']}.pdf",'PDF (*.pdf)')
            if not path:return
        def got(content):
            if action=='preview':PdfDialog(content,'السند #'+str(receipt['serial']),self).exec()
            elif action=='print':self.print_content(receipt['serial'],content,printer)
            else:
                try:Path(path).write_bytes(content);self.statusBar().showMessage('حُفظ PDF.',8000)
                except OSError as e:self.error(str(e))
        self.tasks.submit(lambda:self.api.call('GET',f"/api/receipts/{receipt['serial']}/pdf",binary=True),got,self.error)
    def receipt_details(self):
        receipt=self.selected_receipt()
        if not receipt:return
        d=QDialog(self);d.setWindowTitle('تفاصيل السند #'+str(receipt['serial']));d.resize(1100,600);v=QVBoxLayout(d);v.addWidget(QLabel('العميل: '+receipt['customer']['name']+'   |   أمين الصندوق: '+receipt.get('operator','')))
        t=table(['الكود','الصنف','كرتون','علبة','حبة','الانتهاء','الحالة','السبب','ملاحظات']);fill(t,[[i['product_code'],i['name'],i['cartons'],i['boxes'],i['pieces'],i['expiry'],i['status'],i.get('reason',''),i.get('notes','')] for i in receipt['items']]);v.addWidget(t)
        if not receipt['items']:v.addWidget(QLabel('سجل مستورد من Excel دون تفاصيل أصناف؛ افتح أو أرفق PDF الأصلي.'))
        if receipt.get('notes'):v.addWidget(QLabel('ملاحظات: '+receipt['notes']))
        v.addWidget(button('إغلاق',d.accept,True));d.exec()
    def export_csv(self):
        path,_=QFileDialog.getSaveFileName(self,'تصدير الأرشيف','الأرشيف.csv','CSV (*.csv)')
        if path:self.tasks.submit(lambda:self.api.call('GET','/api/receipts/export.csv?'+urlencode(self.archive_params()),binary=True),lambda content:self.write_file(path,content),self.error)
    def write_file(self,path,content):
        try:Path(path).write_bytes(content);self.statusBar().showMessage('حُفظ الملف: '+path,12000)
        except OSError as e:self.error(str(e))
    def attach_pdf(self):
        if self.identity['role']!='admin':self.error('إرفاق الملفات القديمة متاح للمسؤول.');return
        receipt=self.selected_receipt()
        if not receipt:return
        path,_=QFileDialog.getOpenFileName(self,'إرفاق PDF الأصلي','','PDF (*.pdf)')
        if not path:return
        try:content=Path(path).read_bytes()
        except OSError as e:self.error(str(e));return
        self.tasks.submit(lambda:self.api.call('POST',f"/api/receipts/{receipt['serial']}/attach-pdf",content=content),lambda _:self.load_archive(),self.error)

    def build_masters(self,kind,title):
        page,v=self.title(title,'بحث سريع بالاسم والكود • إضافة سجلات جديدة • السندات السابقة تحتفظ ببياناتها الأصلية')
        q=QLineEdit();q.setPlaceholderText('بحث بالاسم أو الكود…');h=QHBoxLayout();h.addWidget(q,1);h.addWidget(button('بحث',lambda:self.load_masters(kind,True)));h.addWidget(button('إضافة جديد',lambda:self.master_dialog(kind)))
        edit=button('تعديل المحدد',lambda:self.master_dialog(kind,True),True);edit.setEnabled(self.identity['role']=='admin');h.addWidget(edit);v.addLayout(h)
        headers=['الكود','الاسم','رقم المندوب','اسم المندوب','الفرع','المدينة','القطاع'] if kind=='customers' else ['الكود','الاسم']
        t=table(headers);v.addWidget(t,1);count=QLabel();v.addWidget(count);h=QHBoxLayout();h.addWidget(button('السابق',lambda:self.master_page(kind,-1),True));h.addWidget(button('التالي',lambda:self.master_page(kind,1),True));h.addStretch();v.addLayout(h)
        self.master_pages[kind]={'query':q,'table':t,'count':count,'rows':[],'offset':0};q.returnPressed.connect(lambda:self.load_masters(kind,True));self.pages.addWidget(page)
    def master_page(self,kind,delta):self.master_pages[kind]['offset']=max(0,self.master_pages[kind]['offset']+delta*100);self.load_masters(kind)
    def load_masters(self,kind,reset=False):
        data=self.master_pages[kind]
        if reset:data['offset']=0
        path='/api/masters/'+kind+'?'+urlencode({'q':data['query'].text(),'offset':data['offset']})
        def loaded(result):
            data['rows']=result['items'];keys=['code','name','rep_code','rep_name','branch','city','sector'] if kind=='customers' else ['code','name'];fill(data['table'],[[r.get(k,'') for k in keys] for r in data['rows']]);data['count'].setText(f"{data['offset']+1 if data['rows'] else 0}–{data['offset']+len(data['rows'])} من {result['total']} سجل")
            if data['rows']:data['table'].selectRow(0)
        self.tasks.submit(lambda:self.api.call('GET',path),loaded,self.error)
    def master_dialog(self,kind,edit=False):
        data=self.master_pages[kind];existing={}
        if edit:
            row=data['table'].currentRow()
            if not 0<=row<len(data['rows']):return
            existing=data['rows'][row]
        d=QDialog(self);d.setWindowTitle('تعديل البيانات' if edit else 'إضافة سجل جديد');d.resize(520,400);v=QVBoxLayout(d);f=QFormLayout();fields={}
        keys=[('code','الكود'),('name','الاسم')]+([('rep_code','رقم المندوب'),('rep_name','اسم المندوب'),('branch','الفرع'),('city','المدينة'),('sector','القطاع')] if kind=='customers' else [])
        for key,label in keys:
            widget=QLineEdit(existing.get(key,''));f.addRow(label,widget);fields[key]=widget
        fields['code'].setReadOnly(edit);v.addLayout(f);buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel);v.addWidget(buttons);buttons.accepted.connect(d.accept);buttons.rejected.connect(d.reject)
        if d.exec()!=QDialog.DialogCode.Accepted:return
        entry={k:w.text().strip() for k,w in fields.items()}
        self.tasks.submit(lambda:self.api.call('PUT' if edit else 'POST','/api/masters/'+kind,entry),lambda _:(self.load_masters(kind),self.statusBar().showMessage('حُفظت البيانات مركزيًا.',10000)),self.error)

    def build_settings(self):
        page,v=self.title('الإعدادات والنسخ الاحتياطي','تُطبق الإعدادات على السندات الجديدة فقط. النسخ المحفوظة تبقى ثابتة.')
        self.settings_form=QWidget();f=QFormLayout(self.settings_form);self.setting_fields={}
        for key,label in [('company','اسم الشركة'),('membership','رقم العضوية'),('document_title','عنوان السند'),('keeper','أمين المستودع'),('accountant','المحاسب')]:
            widget=QLineEdit();f.addRow(label,widget);self.setting_fields[key]=widget
        f.addRow(QLabel('أمين الصندوق هو اسم المستخدم المسجّل دخوله؛ لا يُدخل يدويًا.'))
        self.thresholds=[];self.labels=[]
        for i in range(4):
            row=QHBoxLayout();label=QLineEdit();self.labels.append(label);row.addWidget(label,2)
            if i<3:spin=QSpinBox();spin.setRange(-10000,36500);self.thresholds.append(spin);row.addWidget(QLabel('حتى'));row.addWidget(spin);row.addWidget(QLabel('يومًا متبقيًا'))
            else:row.addWidget(QLabel('أكثر من الحد الثالث'))
            f.addRow('التصنيف '+str(i+1),row)
        self.next_serial=QSpinBox();self.next_serial.setRange(1,2147483646);f.addRow('رقم التسلسل التالي',self.next_serial)
        self.reasons=QTextEdit();self.reasons.setMaximumHeight(95);f.addRow('أسباب الإرجاع (سبب بكل سطر)',self.reasons);v.addWidget(self.settings_form)
        h=QHBoxLayout();self.settings_save=button('حفظ الإعدادات',self.save_settings);h.addWidget(self.settings_save);h.addWidget(button('تحديث من الخادم',self.reload_settings,True));h.addStretch();v.addLayout(h)
        h=QHBoxLayout();backup=button('تنزيل نسخة احتياطية كاملة',self.backup,True);imp=button('استيراد مصنف Excel',self.import_workbook,True);audit=button('سجل العمليات',self.show_audit,True);h.addWidget(backup);h.addWidget(imp);h.addWidget(audit);v.addLayout(h);v.addStretch()
        admin=self.identity['role']=='admin';self.settings_form.setEnabled(admin);self.settings_save.setEnabled(admin);backup.setEnabled(admin);imp.setEnabled(admin);audit.setEnabled(admin);self.pages.addWidget(page)
    def populate_settings(self):
        if not self.settings:return
        for k,w in self.setting_fields.items():w.setText(self.settings.get(k,''))
        for i,w in enumerate(self.thresholds):w.setValue(self.settings['thresholds'][i])
        for i,w in enumerate(self.labels):w.setText(self.settings['labels'][i])
        self.next_serial.setValue(self.settings['next_serial']);self.reasons.setPlainText('\n'.join(self.settings['reasons']))
    def reload_settings(self):
        def done(value):self.settings=value;self.populate_settings();self.render_lines();self.serial_label.setText('التسلسل التالي: '+str(value['next_serial']))
        self.tasks.submit(lambda:self.api.call('GET','/api/settings'),done,self.error)
    def save_settings(self):
        value={k:w.text().strip() for k,w in self.setting_fields.items()};value.update(thresholds=[w.value() for w in self.thresholds],labels=[w.text().strip() for w in self.labels],next_serial=self.next_serial.value(),reasons=[s.strip() for s in self.reasons.toPlainText().splitlines() if s.strip()])
        self.tasks.submit(lambda:self.api.call('PUT','/api/settings',value),lambda _:self.reload_settings(),self.error)
    def backup(self):
        path,_=QFileDialog.getSaveFileName(self,'نسخة احتياطية كاملة','مرتجعات-'+date.today().isoformat()+'.sqlite','Returns backup (*.sqlite)')
        if path:self.tasks.submit(lambda:self.api.call('GET','/api/backup',binary=True),lambda data:self.write_file(path,data),self.error)
    def import_workbook(self):
        path,_=QFileDialog.getOpenFileName(self,'استيراد البيانات والأرشيف','','Excel (*.xlsx *.xlsm)')
        if not path:return
        if QMessageBox.question(self,'استيراد المصنف','ستُضاف/تُحدّث بيانات الأكواد الموجودة وتُحفظ سجلات الأرشيف دون استبدال السندات. تُنشأ نسخة احتياطية تلقائية قبل الاستيراد. متابعة؟')!=QMessageBox.StandardButton.Yes:return
        try:data=Path(path).read_bytes()
        except OSError as e:self.error(str(e));return
        def done(result):self.reload_settings();QMessageBox.information(self,'اكتمل الاستيراد','نتيجة الاستيراد:\n'+json.dumps(result,ensure_ascii=False,indent=2))
        self.tasks.submit(lambda:self.api.call('POST','/api/import-workbook',content=data),done,self.error)
    def show_audit(self):
        def done(rows):
            d=QDialog(self);d.setWindowTitle('سجل العمليات');d.resize(1000,600);v=QVBoxLayout(d);t=table(['الوقت','العملية','التفاصيل']);fill(t,[[r['created_at'],r['action'],r['data']] for r in rows]);v.addWidget(t);d.exec()
        self.tasks.submit(lambda:self.api.call('GET','/api/audit'),done,self.error)

    def build_users(self):
        page,v=self.title('المستخدمون وأمناء الصندوق','اسم المستخدم المسجّل دخوله يُثبت في السند بوصفه أمين الصندوق؛ لا يتغيّر في السندات السابقة.')
        h=QHBoxLayout();add=button('إنشاء مستخدم جديد',self.new_user);edit=button('تعديل المستخدم / كلمة المرور',self.edit_user,True);h.addWidget(add);h.addWidget(edit);h.addWidget(button('تحديث',self.load_users,True));h.addStretch();v.addLayout(h)
        self.users_grid=table(['اسم الدخول','اسم أمين الصندوق','الصلاحية','الحالة']);v.addWidget(self.users_grid,1);self.user_rows=[];admin=self.identity['role']=='admin';add.setEnabled(admin);edit.setEnabled(admin)
        if not admin:v.addWidget(QLabel('إدارة الحسابات متاحة للمسؤول فقط.'))
        self.pages.addWidget(page)
    def load_users(self):
        if self.identity['role']!='admin':return
        def done(rows):self.user_rows=rows;fill(self.users_grid,[[u['username'],u['display_name'],'مسؤول' if u['role']=='admin' else 'مستخدم','نشط' if u['active'] else 'معطل'] for u in rows])
        self.tasks.submit(lambda:self.api.call('GET','/api/users'),done,self.error)
    def user_dialog(self,existing=None):
        d=QDialog(self);d.setWindowTitle('تعديل المستخدم' if existing else 'إنشاء مستخدم');d.resize(480,360);v=QVBoxLayout(d);f=QFormLayout();username=QLineEdit((existing or {}).get('username',''));username.setReadOnly(bool(existing));name=QLineEdit((existing or {}).get('display_name',''));password=QLineEdit();password.setEchoMode(QLineEdit.EchoMode.Password);role=QComboBox();role.addItem('مستخدم','staff');role.addItem('مسؤول','admin');role.setCurrentIndex(1 if (existing or {}).get('role')=='admin' else 0);active=QCheckBox('الحساب نشط');active.setChecked(bool((existing or {}).get('active',True)))
        for title,w in [('اسم الدخول',username),('اسم أمين الصندوق',name),('كلمة مرور جديدة (اختيارية)' if existing else 'كلمة المرور (8 أحرف على الأقل)',password),('الصلاحية',role)]:f.addRow(title,w)
        if existing:f.addRow(active)
        v.addLayout(f);v.addWidget(QLabel('تظهر هوية المستخدم في كل سند يصدره. تعديل الحساب يلغي جلساته الحالية.'));b=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel);b.accepted.connect(d.accept);b.rejected.connect(d.reject);v.addWidget(b)
        if d.exec()!=QDialog.DialogCode.Accepted:return None
        return {'username':username.text().strip(),'display_name':name.text().strip(),'password':password.text(),'role':role.currentData(),'active':active.isChecked()}
    def new_user(self):
        data=self.user_dialog()
        if data:
            data.pop('active');self.tasks.submit(lambda:self.api.call('POST','/api/users',data),lambda _:self.load_users(),self.error)
    def edit_user(self):
        row=self.users_grid.currentRow()
        if not 0<=row<len(self.user_rows):return
        data=self.user_dialog(self.user_rows[row])
        if data:
            username=data.pop('username');data['new_password']=data.pop('password');self.tasks.submit(lambda:self.api.call('PUT','/api/users/'+username,data),lambda _:self.load_users(),self.error)
    def logout(self):
        self.save_draft();self.want_logout=True
        self.tasks.submit(lambda:self.api.call('POST','/auth/logout'),lambda _:self.close(),lambda _:self.close())
    def closeEvent(self,event):
        if hasattr(self,'autosave'):self.autosave.stop();self.save_draft()
        # All requests use durable idempotency IDs; an interrupted connection can be retried.
        self.tasks.pool.waitForDone(2000)
        event.accept()


def main():
    app=QApplication(sys.argv);app.setApplicationName('ReturnsManager');app.setOrganizationName('Shanifi');app.setLayoutDirection(Qt.LayoutDirection.RightToLeft);app.setStyle('Fusion');app.setStyleSheet(STYLE);app.setFont(QFont('Segoe UI',10))
    while True:
        login=LoginDialog()
        if login.exec()!=QDialog.DialogCode.Accepted:return 0
        window=MainWindow(login.api,login.identity);window.show();app.exec()
        if not getattr(window,'want_logout',False):return 0

if __name__=='__main__':raise SystemExit(main())
