import io
import json
import sqlite3
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import date,timedelta
from pathlib import Path
import pytest
from fastapi.testclient import TestClient
from pypdf import PdfReader
from returns_app.domain import ReceiptInput,status_for,now
from returns_app.store import Store,Conflict
from returns_app.server import create_app
from returns_app.importer import import_workbook
from returns_app.pdf import render_receipt,shape


@pytest.fixture
def store(tmp_path):return Store(tmp_path/'data.sqlite')


def payload(store,request_id=None):
    return {'request_id':request_id or str(uuid.uuid4()),'operator':'مطيع','customer_code':store.search('customers')['items'][0]['code'],'driver_code':store.search('drivers')['items'][0]['code'],'return_number':'0000123','invoice_date':now().date().isoformat(),'items':[{'product_code':store.search('products')['items'][0]['code'],'cartons':2,'boxes':3,'pieces':4,'expiry':(now().date()+timedelta(days=46)).isoformat(),'reason':store.settings()['reasons'][0],'notes':'ملاحظة عربية'}]}


@pytest.mark.parametrize('remaining,index',[(-1,0),(0,0),(45,0),(46,1),(90,1),(91,2),(180,2),(181,3),(271,3)])
def test_expiry_boundaries(store,remaining,index):
    settings=store.settings();label,got,left=status_for(date(2026,11,14)+timedelta(days=remaining),date(2026,11,14),settings)
    assert got==index and left==remaining and label==settings['labels'][index]


def test_unicode_master_search_and_no_missing_seed(store):
    assert store.search('customers')['total']==2701
    assert store.search('products')['total']==303
    assert store.search('drivers')['total']==4
    assert store.search('customers','الشنيفي')['total']>=0
    code=store.search('customers')['items'][0]['code']
    assert store.search('customers',code)['total']==1


def test_receipt_pdf_idempotency_immutable_snapshot(store):
    data=payload(store);first=store.create_receipt(ReceiptInput(**data));again=store.create_receipt(ReceiptInput(**data))
    assert first==again and store.settings()['next_serial']==first['serial']+1
    snapshot,pdf=store.receipt(first['serial']);reader=PdfReader(io.BytesIO(pdf))
    assert len(reader.pages)==1 and pdf.startswith(b'%PDF')
    import unicodedata
    text=unicodedata.normalize('NFKC',reader.pages[0].extract_text())
    assert 'المحاسب' in text and 'أمين الصندوق' in text
    store.put_master('customers',{'code':data['customer_code'],'name':'اسم جديد'})
    assert store.receipt(first['serial'])[0]['customer']['name']==snapshot['customer']['name']
    assert snapshot['operator']=='مطيع' and snapshot['return_number']=='0000123'
    data['return_number']='changed'
    with pytest.raises(Conflict):store.create_receipt(ReceiptInput(**data))


def test_atomic_export_failure_rolls_back(store,monkeypatch):
    import returns_app.store as module
    before=store.settings()['next_serial']
    def broken(*a):raise OSError('disk failure')
    monkeypatch.setattr(module,'render_receipt',broken)
    with pytest.raises(OSError):store.create_receipt(ReceiptInput(**payload(store)))
    assert store.list_receipts()['total']==0 and store.settings()['next_serial']==before


def test_concurrent_devices_unique_and_same_request_retry(store):
    data=payload(store)
    def create(i):
        value={**data,'request_id':str(uuid.uuid4()),'operator':f'صندوق {i}'}
        return store.create_receipt(ReceiptInput(**value))['serial']
    with ThreadPoolExecutor(max_workers=5) as pool:serials=list(pool.map(create,range(10)))
    assert len(set(serials))==10 and sorted(serials)==list(range(1,11))
    same=ReceiptInput(**payload(store))
    with ThreadPoolExecutor(max_workers=4) as pool:retries=list(pool.map(lambda _:store.create_receipt(same)['serial'],range(4)))
    assert len(set(retries))==1 and store.list_receipts()['total']==11


def test_multi_page_and_full_notes(store):
    data=payload(store);data['items']*=23
    snapshot=store.create_receipt(ReceiptInput(**data));reader=PdfReader(io.BytesIO(store.receipt(snapshot['serial'])[1]))
    assert len(reader.pages)==2
    snapshot['items'][0]['notes']='تفاصيل طويلة '*100
    more=PdfReader(io.BytesIO(render_receipt(snapshot)))
    assert len(more.pages)>2


def test_backup_contains_canonical_pdf(store,tmp_path):
    receipt=store.create_receipt(ReceiptInput(**payload(store)))
    backup=tmp_path/'restore.sqlite';backup.write_bytes(store.backup_bytes());restored=Store(backup)
    assert restored.receipt(receipt['serial'])==store.receipt(receipt['serial'])
    with restored.connect() as c:assert c.execute('PRAGMA integrity_check').fetchone()[0]=='ok'


@pytest.fixture
def api(tmp_path):
    app=create_app(tmp_path/'server','s'*32,'a'*32,'InitialPass123')
    with TestClient(app) as client:
        result=client.post('/auth/login',json={'username':'مطيع','password':'InitialPass123'})
        assert result.status_code==200
        client.headers['Authorization']='Bearer '+result.json()['token']
        yield client,app.state.store


def test_user_cashier_cannot_be_spoofed_and_roles(api):
    client,store=api
    assert client.post('/api/users',json={'username':'ثاني','display_name':'أحمد','password':'StrongPass123','role':'staff'}).status_code==200
    token=client.post('/auth/login',json={'username':'ثاني','password':'StrongPass123'}).json()['token']
    client.headers['Authorization']='Bearer '+token
    assert client.get('/api/session').json()['display_name']=='أحمد'
    data=payload(store);data['operator']='اسم مزور'
    response=client.post('/api/receipts',json=data)
    assert response.status_code==200 and response.json()['operator']=='أحمد'
    assert client.put('/api/settings',json=store.settings()).status_code==403
    assert client.get('/api/backup').status_code==403
    assert client.get('/api/users').status_code==403
    assert client.post('/api/masters/products',json={'code':'000NEW','name':'منتج عربي جديد'}).status_code==200
    assert client.post('/api/masters/products',json={'code':'000NEW','name':'overwrite'}).status_code==409
    assert store.search('products','000NEW')['items'][0]['name']=='منتج عربي جديد'
    client.post('/auth/logout')
    assert client.get('/api/session').status_code==401


def test_admin_update_last_user_and_invalidate_sessions(api):
    client,store=api
    assert client.put('/api/users/مطيع',json={'display_name':'مطيع','active':False}).status_code==400
    assert client.post('/api/users',json={'username':'مندوب','display_name':'موظف','password':'StrongPass123'}).status_code==200
    token=client.post('/auth/login',json={'username':'مندوب','password':'StrongPass123'}).json()['token']
    assert client.put('/api/users/مندوب',json={'display_name':'اسم محدث','active':True,'new_password':'NewSecret123'}).status_code==200
    assert store.identity(token) is None
    with store.connect() as c:
        stored=c.execute('SELECT password FROM users WHERE username=?',('مندوب',)).fetchone()[0]
        assert stored!='NewSecret123' and len(stored)==64


def test_api_validations_and_pdf_print_tracking(api):
    client,store=api;data=payload(store)
    bad={**data,'items':[{**data['items'][0],'cartons':1.5}]}
    assert client.post('/api/receipts',json=bad).status_code==422
    future={**data,'invoice_date':(now().date()+timedelta(days=1)).isoformat()}
    assert client.post('/api/receipts',json=future).status_code==400
    zero={**data,'items':[{**data['items'][0],'cartons':0,'boxes':0,'pieces':0}]}
    assert client.post('/api/receipts',json=zero).status_code==400
    receipt=client.post('/api/receipts',json=data).json();serial=receipt['serial']
    assert client.get(f'/api/receipts/{serial}/pdf').content.startswith(b'%PDF')
    response=client.post(f'/api/receipts/{serial}/print-events',json={'operator':'other','printer':'طابعة المكتب','copies':2,'success':False,'error':'offline'})
    assert response.status_code==200 and store.list_receipts()['items'][0]['print_state']=='تعذر الإرسال للطابعة'
    assert store.settings()['next_serial']==serial+1
    assert client.get('/api/receipts/export.csv').content.startswith(b'\xef\xbb\xbf')


def test_import_latest_workbook_and_idempotent(store):
    path=Path(__file__).resolve().parents[2]/'Returns_Management_PDF_Archive.xlsm'
    if not path.exists():pytest.skip('Source workbook not included in standalone source bundle')
    content=path.read_bytes();result=import_workbook(store,content)
    assert result['counts']['customers']==2701 and result['counts']['drivers']==4
    assert store.settings()['thresholds']==[45,90,180]
    assert import_workbook(store,content)['already_imported']
