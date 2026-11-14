from __future__ import annotations
import csv
import hashlib
import io
import json
import sqlite3
import tempfile
from contextlib import contextmanager
from pathlib import Path
from datetime import date
from .domain import now, normalized, status_for, ReceiptInput, SettingsInput
from .pdf import render_receipt, ASSETS


def dump(value):
    return json.dumps(value, ensure_ascii=False, separators=(',', ':'), sort_keys=True)


class Conflict(ValueError):
    pass


class Store:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as c:
            c.executescript('''
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS settings(id INTEGER PRIMARY KEY CHECK(id=1), data TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS masters(kind TEXT NOT NULL, code TEXT NOT NULL, data TEXT NOT NULL, search TEXT NOT NULL, PRIMARY KEY(kind,code));
                CREATE TABLE IF NOT EXISTS receipts(serial INTEGER PRIMARY KEY, request_id TEXT UNIQUE NOT NULL, hash TEXT NOT NULL, created_at TEXT NOT NULL, data TEXT NOT NULL, pdf BLOB, print_state TEXT NOT NULL DEFAULT 'محفوظ', legacy INTEGER NOT NULL DEFAULT 0);
                CREATE TABLE IF NOT EXISTS print_events(id INTEGER PRIMARY KEY, serial INTEGER NOT NULL REFERENCES receipts(serial), created_at TEXT NOT NULL, operator TEXT NOT NULL, printer TEXT NOT NULL, copies INTEGER NOT NULL, success INTEGER NOT NULL, error TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, action TEXT NOT NULL, data TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS import_log(source_hash TEXT PRIMARY KEY, created_at TEXT NOT NULL);
            ''')
            if not c.execute('SELECT 1 FROM settings').fetchone():
                seed = json.loads((ASSETS/'seed.json').read_text(encoding='utf-8'))
                self._replace_masters(c, seed)
                c.execute('INSERT INTO settings VALUES(1,?)', (dump(seed['settings']),))

    @contextmanager
    def connect(self):
        c = sqlite3.connect(self.path, timeout=45)
        c.row_factory = sqlite3.Row
        c.execute('PRAGMA foreign_keys=ON')
        try:
            with c:
                yield c
        finally:
            c.close()

    @staticmethod
    def _replace_masters(c, bundle):
        for kind in ['customers','products','drivers']:
            for entry in bundle.get(kind, []):
                Store._put(c, kind, entry)

    @staticmethod
    def _put(c, kind, entry):
        code, name = str(entry.get('code','')).strip(), str(entry.get('name','')).strip()
        if not code or not name or len(code)>100 or len(name)>1000:
            raise ValueError('الكود والاسم مطلوبان وبطول صالح.')
        allowed = {'code','name','rep_code','rep_name','branch','city','sector'} if kind=='customers' else {'code','name'}
        entry = {k: str(v).strip() for k,v in entry.items() if k in allowed}
        entry.update(code=code, name=name)
        c.execute('INSERT INTO masters VALUES(?,?,?,?) ON CONFLICT(kind,code) DO UPDATE SET data=excluded.data,search=excluded.search', (kind,code,dump(entry), normalized(' '.join(entry.values()))))

    def settings(self):
        with self.connect() as c:
            return json.loads(c.execute('SELECT data FROM settings WHERE id=1').fetchone()[0])

    def set_settings(self, value: SettingsInput):
        settings = value.model_dump()
        if sorted(settings['thresholds']) != settings['thresholds'] or len(set(settings['thresholds']))!=3:
            raise ValueError('حدود التصنيف يجب أن تكون ثلاثة أعداد متزايدة.')
        if any(not s.strip() for s in settings['labels']) or len(set(settings['labels']))!=4:
            raise ValueError('أدخل أربع مسميات مختلفة وغير فارغة.')
        with self.connect() as c:
            c.execute('BEGIN IMMEDIATE')
            maximum = c.execute('SELECT coalesce(max(serial),0) FROM receipts').fetchone()[0]
            if settings['next_serial'] <= maximum:
                raise ValueError('الرقم التالي يجب أن يكون أكبر من جميع السندات المحفوظة.')
            c.execute('UPDATE settings SET data=? WHERE id=1',(dump(settings),))
            self.audit(c,'settings',settings)
        return settings

    @staticmethod
    def audit(c, action, data):
        c.execute('INSERT INTO audit(created_at,action,data) VALUES(?,?,?)',(now().isoformat(),action,dump(data)))

    def search(self, kind, query='', limit=100, offset=0):
        if kind not in ['customers','products','drivers']:
            raise ValueError('نوع بيانات غير صالح.')
        tokens = normalized(query).split()
        where = 'kind=?' + ''.join(" AND instr(search,?)>0" for _ in tokens)
        params = [kind]+tokens
        with self.connect() as c:
            total = c.execute('SELECT count(*) FROM masters WHERE '+where,params).fetchone()[0]
            rows = c.execute('SELECT data FROM masters WHERE '+where+' ORDER BY code LIMIT ? OFFSET ?',params+[min(limit,200),max(offset,0)]).fetchall()
            return {'total':total, 'items':[json.loads(r[0]) for r in rows]}

    def put_master(self,kind,entry):
        if kind not in ['customers','products','drivers']:
            raise ValueError('نوع بيانات غير صالح.')
        with self.connect() as c:
            self._put(c,kind,entry)
            self.audit(c,'master:'+kind,entry)

    def create_receipt(self, request: ReceiptInput):
        payload = request.model_dump(mode='json')
        digest = hashlib.sha256(dump({k:v for k,v in payload.items() if k!='operator'}).encode()).hexdigest()
        # Sequence, immutable snapshot and actual PDF commit together under one writer lock.
        with self.connect() as c:
            c.execute('BEGIN IMMEDIATE')
            existing = c.execute('SELECT hash,data FROM receipts WHERE request_id=?',(request.request_id,)).fetchone()
            if existing:
                if existing['hash'] != digest:
                    raise Conflict('معرّف العملية مستخدم لمدخلات مختلفة؛ أنشئ سندًا جديدًا.')
                return json.loads(existing['data'])
            settings = json.loads(c.execute('SELECT data FROM settings WHERE id=1').fetchone()[0])
            def lookup(kind,code):
                if not code and kind=='drivers':
                    return {'code':'','name':''}
                row = c.execute('SELECT data FROM masters WHERE kind=? AND code=?',(kind,code)).fetchone()
                if not row:
                    raise ValueError('الكود غير موجود في قاعدة البيانات: '+code)
                return json.loads(row[0])
            stamp = now()
            if request.invoice_date and request.invoice_date > stamp.date():
                raise ValueError('تاريخ الفاتورة لا يمكن أن يكون مستقبليًا.')
            customer = lookup('customers',request.customer_code)
            driver = lookup('drivers',request.driver_code)
            items = []
            for i,item in enumerate(request.items,1):
                if item.cartons+item.boxes+item.pieces<=0:
                    raise ValueError(f'الصف {i}: أدخل كمية أكبر من صفر.')
                if item.reason and item.reason not in settings['reasons']:
                    raise ValueError(f'الصف {i}: اختر سببًا معتمدًا من القائمة.')
                if item.expiry.year < 1900:
                    raise ValueError('أدخل تاريخ انتهاء صحيحًا.')
                product = lookup('products',item.product_code)
                label,index,remaining = status_for(item.expiry,stamp.date(),settings)
                items.append({**item.model_dump(mode='json'),'name':product['name'],'status':label,'status_index':index,'remaining_days':remaining})
            maximum = c.execute('SELECT coalesce(max(serial),0) FROM receipts').fetchone()[0]
            serial = max(settings['next_serial'],maximum+1)
            snapshot = {**payload,'serial':serial,'created_at':stamp.isoformat(),'customer':customer,'driver':driver,'settings':settings.copy(),'items':items,'totals':{k:sum(i[k] for i in items) for k in ['cartons','boxes','pieces']}}
            pdf = render_receipt(snapshot)
            if not pdf.startswith(b'%PDF') or len(pdf)<100:
                raise ValueError('تعذر إنشاء PDF؛ لم يُعتمد السند.')
            c.execute('INSERT INTO receipts(serial,request_id,hash,created_at,data,pdf) VALUES(?,?,?,?,?,?)',(serial,request.request_id,digest,stamp.isoformat(),dump(snapshot),pdf))
            settings['next_serial']=serial+1
            c.execute('UPDATE settings SET data=? WHERE id=1',(dump(settings),))
            self.audit(c,'receipt',{'serial':serial,'operator':request.operator})
            return snapshot

    def list_receipts(self, query='', start='', end='', offset=0, limit=100):
        with self.connect() as c:
            where,params = [],[]
            if query:
                # JSON contains full immutable customer/driver/item fields.
                where.append('(instr(data,?)>0 OR cast(serial as text)=?)')
                params += [query,query]
            if start:
                where.append('substr(created_at,1,10)>=?');params.append(start)
            if end:
                where.append('substr(created_at,1,10)<=?');params.append(end)
            where=' WHERE '+' AND '.join(where) if where else ''
            total=c.execute('SELECT count(*) FROM receipts'+where,params).fetchone()[0]
            rows=c.execute('SELECT serial,data,print_state,legacy,pdf IS NOT NULL as has_pdf FROM receipts'+where+' ORDER BY serial DESC LIMIT ? OFFSET ?',params+[min(limit,200),max(offset,0)]).fetchall()
            return {'total':total,'items':[{**json.loads(r['data']),'print_state':r['print_state'],'legacy':bool(r['legacy']),'has_pdf':bool(r['has_pdf'])} for r in rows]}

    def receipt(self,serial):
        with self.connect() as c:
            row=c.execute('SELECT data,pdf FROM receipts WHERE serial=?',(serial,)).fetchone()
            if not row:
                raise KeyError('السند غير موجود.')
            return json.loads(row['data']),row['pdf']

    def print_event(self,serial,operator,printer,copies,success,error=''):
        with self.connect() as c:
            c.execute('INSERT INTO print_events(serial,created_at,operator,printer,copies,success,error) VALUES(?,?,?,?,?,?,?)',(serial,now().isoformat(),operator,printer,copies,int(success),error))
            c.execute('UPDATE receipts SET print_state=? WHERE serial=?',('أُرسل للطابعة' if success else 'تعذر الإرسال للطابعة',serial))

    def csv_export(self, query='', start='', end=''):
        buf=io.StringIO(newline='');w=csv.writer(buf)
        w.writerow(['التسلسل','التاريخ','كود العميل','العميل','عدد الأصناف','كرتون','علبة','حبة','رقم الإرجاع','تاريخ الفاتورة','رقم السائق','السائق','أمين الصندوق','المخازن','حالة الطباعة'])
        offset=0
        while True:
            result=self.list_receipts(query,start,end,offset,200)
            for r in result['items']:
                row=[r['serial'],r['created_at'],r['customer']['code'],r['customer']['name'],len(r['items']) if not r.get('legacy_count') else r['legacy_count'],*[r['totals'][k] for k in ['cartons','boxes','pieces']],r.get('return_number',''),r.get('invoice_date',''),r['driver']['code'],r['driver']['name'],r.get('operator',''),'، '.join(sorted(set(i['status'] for i in r['items']))),r['print_state']]
                w.writerow([("'"+str(v) if isinstance(v,str) and v.startswith(('=','+','-','@','\t')) else v) for v in row])
            offset+=len(result['items'])
            if offset>=result['total']:break
        return ('\ufeff'+buf.getvalue()).encode('utf-8')

    def backup_bytes(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'backup.sqlite'
            with self.connect() as source,sqlite3.connect(path) as target:
                source.backup(target)
            return path.read_bytes()

    def setup_users(self, bootstrap_password):
        with self.connect() as c:
            c.executescript('''
                CREATE TABLE IF NOT EXISTS users(username TEXT PRIMARY KEY, display_name TEXT NOT NULL, salt TEXT NOT NULL, password TEXT NOT NULL, role TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1);
                CREATE TABLE IF NOT EXISTS sessions(token_hash TEXT PRIMARY KEY, username TEXT NOT NULL REFERENCES users(username), expires TEXT NOT NULL);
            ''')
            empty=not c.execute('SELECT 1 FROM users').fetchone()
        if empty and bootstrap_password:
            self.create_user('مطيع','مطيع',bootstrap_password,'admin')

    def create_user(self,username,display_name,password,role='staff'):
        import secrets
        username=username.strip();display_name=display_name.strip()
        if not username or not display_name or len(username)>100 or len(display_name)>200:raise ValueError('اسم الدخول واسم أمين الصندوق مطلوبان.')
        if len(password)<8 or len(password)>200:raise ValueError('كلمة المرور من 8 إلى 200 حرف.')
        if role not in ['admin','staff']:raise ValueError('صلاحية المستخدم غير صالحة.')
        salt=secrets.token_hex(16)
        hashed=hashlib.pbkdf2_hmac('sha256',password.encode(),bytes.fromhex(salt),300000).hex()
        with self.connect() as c:
            if c.execute('SELECT 1 FROM users WHERE username=?',(username,)).fetchone():raise Conflict('اسم الدخول موجود بالفعل.')
            c.execute('INSERT INTO users(username,display_name,salt,password,role) VALUES(?,?,?,?,?)',(username,display_name,salt,hashed,role))
            self.audit(c,'create_user',{'username':username,'display_name':display_name,'role':role})

    def login(self,username,password):
        import secrets,hmac
        from datetime import timedelta
        with self.connect() as c:
            row=c.execute('SELECT * FROM users WHERE username=? AND active=1',(username,)).fetchone()
            salt=row['salt'] if row else '0'*32
            hashed=hashlib.pbkdf2_hmac('sha256',password.encode(),bytes.fromhex(salt),300000).hex()
            if not row or not hmac.compare_digest(hashed,row['password']):raise PermissionError('اسم المستخدم أو كلمة المرور غير صحيحة.')
            token=secrets.token_urlsafe(40)
            c.execute('DELETE FROM sessions WHERE expires<?',(now().isoformat(),))
            c.execute('INSERT INTO sessions VALUES(?,?,?)',(hashlib.sha256(token.encode()).hexdigest(),username,(now()+timedelta(hours=12)).isoformat()))
            self.audit(c,'login',{'username':username})
            return token

    def identity(self,token):
        with self.connect() as c:
            row=c.execute('SELECT u.username,u.display_name,u.role FROM sessions s JOIN users u ON u.username=s.username WHERE token_hash=? AND expires>? AND u.active=1',(hashlib.sha256(token.encode()).hexdigest(),now().isoformat())).fetchone()
            return dict(row) if row else None

    def list_users(self):
        with self.connect() as c:return [dict(r) for r in c.execute('SELECT username,display_name,role,active FROM users ORDER BY username')]

    def update_user(self,username,display_name,active,new_password='',role=None):
        import secrets
        with self.connect() as c:
            c.execute('BEGIN IMMEDIATE')
            row=c.execute('SELECT * FROM users WHERE username=?',(username,)).fetchone()
            if not row:raise KeyError('المستخدم غير موجود.')
            next_role=role or row['role']
            if next_role not in ['admin','staff'] or not display_name.strip():raise ValueError('أدخل اسمًا وصلاحية صحيحين.')
            if row['role']=='admin' and (not active or next_role!='admin') and c.execute("SELECT count(*) FROM users WHERE role='admin' AND active=1").fetchone()[0]<=1:raise ValueError('لا يمكن تعطيل آخر مسؤول.')
            c.execute('UPDATE users SET display_name=?,active=?,role=? WHERE username=?',(display_name.strip(),int(active),next_role,username))
            if new_password:
                if len(new_password)<8:raise ValueError('كلمة المرور الجديدة 8 أحرف على الأقل.')
                salt=secrets.token_hex(16);hashed=hashlib.pbkdf2_hmac('sha256',new_password.encode(),bytes.fromhex(salt),300000).hex()
                c.execute('UPDATE users SET salt=?,password=? WHERE username=?',(salt,hashed,username))
            c.execute('DELETE FROM sessions WHERE username=?',(username,))
            self.audit(c,'update_user',{'username':username,'active':active,'role':next_role})
