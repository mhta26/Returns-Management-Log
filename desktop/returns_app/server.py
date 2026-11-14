import argparse
import hmac
import json
import os
import secrets
from contextlib import closing
from pathlib import Path
from typing import Annotated
from fastapi import FastAPI, Header, Depends, HTTPException, Request
from fastapi.responses import Response, JSONResponse
from pydantic import BaseModel, Field
from .store import Store, Conflict
from .domain import ReceiptInput, SettingsInput, now
from .importer import import_workbook


class PrintEvent(BaseModel):
    operator: str = Field(min_length=1,max_length=200)
    printer: str = Field(max_length=500)
    copies: int = Field(ge=1,le=999)
    success: bool
    error: str = Field(default='',max_length=2000)


def create_app(data_dir:Path,staff_token:str,admin_token:str,bootstrap_password:str|None=None):
    if len(staff_token)<16 or len(admin_token)<16:raise ValueError('استخدم مفاتيح اتصال قوية.')
    data_dir=Path(data_dir);store=Store(data_dir/'returns.sqlite')
    app=FastAPI(title='Returns Manager',docs_url=None,redoc_url=None,openapi_url=None)
    app.state.store=store
    store.setup_users(bootstrap_password)

    def auth(authorization:Annotated[str|None,Header()]=None):
        supplied=(authorization or '').removeprefix('Bearer ')
        identity=store.identity(supplied)
        if identity:return identity
        raise HTTPException(401,'انتهت جلسة الدخول أو بيانات الاتصال غير صحيحة؛ سجل الدخول مجددًا.')

    def admin(identity=Depends(auth)):
        if identity['role']!='admin':raise HTTPException(403,'هذه العملية متاحة للمسؤول فقط.')

    @app.exception_handler(ValueError)
    async def invalid(_,exc):return JSONResponse({'detail':str(exc)},status_code=409 if isinstance(exc,Conflict) else 400)

    @app.exception_handler(KeyError)
    async def missing(_,exc):return JSONResponse({'detail':str(exc)},status_code=404)

    from collections import defaultdict
    import time, threading
    attempts=defaultdict(list);attempt_lock=threading.Lock()

    class Login(BaseModel):
        username:str=Field(min_length=1,max_length=100)
        password:str=Field(min_length=1,max_length=200)

    class UserInput(BaseModel):
        username:str=Field(min_length=1,max_length=100)
        display_name:str=Field(min_length=1,max_length=200)
        password:str=Field(min_length=8,max_length=200)
        role:str='staff'

    class UserUpdate(BaseModel):
        display_name:str=Field(min_length=1,max_length=200)
        active:bool=True
        new_password:str=Field(default='',max_length=200)
        role:str|None=None

    @app.post('/auth/login')
    def login(value:Login,request:Request):
        address=request.client.host if request.client else 'local'
        with attempt_lock:
            recent=[t for t in attempts[address] if time.monotonic()-t<60]
            if len(recent)>=8:raise HTTPException(429,'محاولات كثيرة؛ انتظر دقيقة ثم أعد المحاولة.')
            attempts[address]=recent+[time.monotonic()]
            if len(attempts)>2048:
                for key in list(attempts):
                    if not attempts[key] or time.monotonic()-max(attempts[key])>60:del attempts[key]
        try:token=store.login(value.username,value.password)
        except PermissionError as e:raise HTTPException(401,str(e))
        return {'token':token,'identity':store.identity(token)}

    @app.post('/auth/logout')
    def logout(authorization:Annotated[str|None,Header()]=None):
        import hashlib
        with store.connect() as c:c.execute('DELETE FROM sessions WHERE token_hash=?',(hashlib.sha256((authorization or '').removeprefix('Bearer ').encode()).hexdigest(),))
        return {'ok':True}

    @app.get('/api/users',dependencies=[Depends(admin)])
    def users():return store.list_users()

    @app.post('/api/users',dependencies=[Depends(admin)])
    def create_user(value:UserInput):store.create_user(**value.model_dump());return {'ok':True}

    @app.put('/api/users/{username}',dependencies=[Depends(admin)])
    def update_user(username:str,value:UserUpdate):store.update_user(username,**value.model_dump());return {'ok':True}

    @app.get('/health')
    def health():return {'status':'ok','version':'1.0.0'}

    @app.get('/api/session')
    def session(identity=Depends(auth)):return {**identity,'server_time':now().isoformat(),'settings':store.settings()}

    @app.get('/api/masters/{kind}',dependencies=[Depends(auth)])
    def masters(kind:str,q:str='',limit:int=100,offset:int=0):return store.search(kind,q,limit,offset)

    @app.post('/api/masters/{kind}',dependencies=[Depends(auth)])
    def add_master(kind:str,entry:dict):
        if kind not in ['customers','products','drivers']:raise ValueError('نوع بيانات غير صالح.')
        with store.connect() as c:
            c.execute('BEGIN IMMEDIATE')
            if c.execute('SELECT 1 FROM masters WHERE kind=? AND code=?',(kind,str(entry.get('code','')).strip())).fetchone():raise Conflict('الكود موجود؛ لا يمكن إنشاء نسخة مكررة منه.')
            store._put(c,kind,entry)
            store.audit(c,'new_master:'+kind,entry)
        return {'ok':True}

    @app.put('/api/masters/{kind}',dependencies=[Depends(admin)])
    def put_master(kind:str,entry:dict):store.put_master(kind,entry);return {'ok':True}

    @app.get('/api/settings',dependencies=[Depends(auth)])
    def settings():return store.settings()

    @app.put('/api/settings',dependencies=[Depends(admin)])
    def set_settings(value:SettingsInput):return store.set_settings(value)

    @app.post('/api/receipts')
    def receipt(value:ReceiptInput,identity=Depends(auth)):
        value.operator=identity['display_name']
        return store.create_receipt(value)

    @app.get('/api/receipts',dependencies=[Depends(auth)])
    def receipts(q:str='',start:str='',end:str='',offset:int=0,limit:int=100):return store.list_receipts(q,start,end,offset,limit)

    @app.get('/api/receipts/export.csv',dependencies=[Depends(auth)])
    def export(q:str='',start:str='',end:str=''):return Response(store.csv_export(q,start,end),media_type='text/csv; charset=utf-8')

    @app.get('/api/receipts/{serial}',dependencies=[Depends(auth)])
    def get_receipt(serial:int):return store.receipt(serial)[0]

    @app.get('/api/receipts/{serial}/pdf',dependencies=[Depends(auth)])
    def pdf(serial:int):
        snapshot,content=store.receipt(serial)
        if not content:
            if not snapshot['items']:raise HTTPException(404,'السجل القديم لا يحتوي PDF أو تفاصيل الأصناف. أرفق PDF الأصلي من الأرشيف.')
            from .pdf import render_receipt
            content=render_receipt(snapshot)
        return Response(content,media_type='application/pdf',headers={'Content-Disposition':f'inline; filename="receipt-{serial}.pdf"'})

    @app.post('/api/receipts/{serial}/print-events')
    def print_event(serial:int,value:PrintEvent,identity=Depends(auth)):
        value.operator=identity['display_name']
        store.receipt(serial)
        store.print_event(serial,**value.model_dump());return {'ok':True}

    async def bounded_body(request,limit):
        parts=[];size=0
        async for part in request.stream():
            size+=len(part)
            if size>limit:raise HTTPException(413,'الملف أكبر من الحجم المسموح.')
            parts.append(part)
        return b''.join(parts)

    @app.post('/api/import-workbook',dependencies=[Depends(admin)])
    async def import_data(request:Request):
        content=await bounded_body(request,20_000_000)
        backup=data_dir/'backups';backup.mkdir(exist_ok=True)
        (backup/('before-import-'+now().strftime('%Y%m%d-%H%M%S-%f')+'.sqlite')).write_bytes(store.backup_bytes())
        result=import_workbook(store,content)
        import hashlib
        originals=data_dir/'imports';originals.mkdir(exist_ok=True)
        (originals/(hashlib.sha256(content).hexdigest()+'.xlsx')).write_bytes(content)
        return result

    @app.post('/api/receipts/{serial}/attach-pdf',dependencies=[Depends(admin)])
    async def attach(serial:int,request:Request):
        store.receipt(serial);content=await bounded_body(request,30_000_000)
        if not content.startswith(b'%PDF-'):raise HTTPException(400,'ملف PDF غير صالح.')
        with store.connect() as c:
            exists=c.execute('SELECT pdf FROM receipts WHERE serial=?',(serial,)).fetchone()[0]
            if exists:raise Conflict('يوجد PDF محفوظ لهذا السند؛ لا يمكن استبداله.')
            c.execute('UPDATE receipts SET pdf=? WHERE serial=?',(content,serial))
            store.audit(c,'attach_pdf',{'serial':serial})
        return {'ok':True}

    @app.get('/api/backup',dependencies=[Depends(admin)])
    def backup():return Response(store.backup_bytes(),media_type='application/octet-stream')

    @app.get('/api/audit',dependencies=[Depends(admin)])
    def audit():
        with store.connect() as c:return [dict(r) for r in c.execute('SELECT * FROM audit ORDER BY id DESC LIMIT 200')]

    return app


def main():
    import sys
    if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8',errors='replace')
    parser=argparse.ArgumentParser(description='خادم إدارة المرتجعات المركزي')
    parser.add_argument('--data-dir',type=Path,default=Path.home()/'Documents'/'ReturnsManagerServer')
    parser.add_argument('--host',default='0.0.0.0');parser.add_argument('--port',type=int,default=8765)
    parser.add_argument('--restore',type=Path,help='استعادة نسخة SQLite قبل تشغيل الخدمة')
    args=parser.parse_args();args.data_dir.mkdir(parents=True,exist_ok=True)
    if args.restore:
        import sqlite3
        with closing(sqlite3.connect(f'file:{args.restore.resolve().as_posix()}?mode=ro',uri=True)) as source:
            if source.execute('PRAGMA integrity_check').fetchone()[0]!='ok':raise SystemExit('النسخة الاحتياطية غير سليمة.')
            if not source.execute("SELECT 1 FROM sqlite_master WHERE name='receipts'").fetchone():raise SystemExit('ليست نسخة إدارة المرتجعات.')
            # Restore only with no existing destination; user keeps current data safely.
            destination=args.data_dir/'returns.sqlite'
            if destination.exists():raise SystemExit('اختر مجلد بيانات جديدًا للاستعادة؛ لم تُستبدل البيانات الحالية.')
            with closing(sqlite3.connect(destination)) as target:
                with target:source.backup(target)
    credentials=args.data_dir/'connection-keys.json'
    if not credentials.exists():
        credentials.write_text(json.dumps({'staff_token':secrets.token_urlsafe(32),'admin_token':secrets.token_urlsafe(32)},indent=2),encoding='utf-8')
        if os.name!='nt':credentials.chmod(0o600)
    keys=json.loads(credentials.read_text(encoding='utf-8'))
    if 'initial_password' not in keys:
        keys['initial_password']=secrets.token_urlsafe(12)
        credentials.write_text(json.dumps(keys,ensure_ascii=False,indent=2),encoding='utf-8')
    print('First administrator username: مطيع (see initial_password in connection-keys.json)')
    print('Returns server: port',args.port)
    print('Connection keys file:',credentials)
    print('Keep this window open. Enter this computer IP in each desktop client.')
    import uvicorn
    uvicorn.run(create_app(args.data_dir,keys['staff_token'],keys['admin_token'],keys['initial_password']),host=args.host,port=args.port,access_log=False)

if __name__=='__main__':main()
