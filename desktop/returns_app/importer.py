"""Read workbook data only. VBA, external links and embedded instructions never execute."""
from __future__ import annotations
from io import BytesIO
import hashlib
import json
from datetime import datetime, date
from zipfile import ZipFile, BadZipFile
import openpyxl
from .store import dump, Conflict
from .domain import now


def code(value):
    if value is None:return ''
    if isinstance(value,float) and value.is_integer():return str(int(value))
    return str(value).strip()


def iso(value):
    if isinstance(value,(datetime,date)):return value.isoformat()[:10]
    try:return date.fromisoformat(str(value)[:10]).isoformat()
    except (ValueError,TypeError):return None


def import_workbook(store,content:bytes):
    digest=hashlib.sha256(content).hexdigest()
    try:
        with ZipFile(BytesIO(content)) as z:
            if sum(i.file_size for i in z.infolist())>80_000_000:
                raise ValueError('المصنف أكبر من حد الاستيراد المسموح.')
    except BadZipFile as e:raise ValueError('اختر ملف XLSX أو XLSM صالحًا.') from e
    wb=openpyxl.load_workbook(BytesIO(content),data_only=True,read_only=True,keep_links=False)
    try:
        for required in ['قاعدة العملاء','دليل الأصناف']:
            if required not in wb.sheetnames:raise ValueError('ورقة مطلوبة غير موجودة: '+required)
        bundle={}
        for kind,name,fields in [('customers','قاعدة العملاء',['code','name','rep_code','rep_name','branch','city','sector']),('products','دليل الأصناف',['code','name']),('drivers','قاعدة السائقين',['code','name'])]:
            bundle[kind]=[]
            seen=set()
            if name not in wb.sheetnames:continue
            s=wb[name]
            for row in s.iter_rows(min_row=2,max_col=len(fields),values_only=True):
                if not any(v is not None for v in row):continue
                entry={key:code(value) for key,value in zip(fields,row)}
                if not entry['code'] or not entry['name']:raise ValueError('صف ناقص الكود أو الاسم في '+name)
                if entry['code'] in seen:raise ValueError('كود مكرر في '+name+': '+entry['code'])
                seen.add(entry['code']);bundle[kind].append(entry)
        with store.connect() as c:
            c.execute('BEGIN IMMEDIATE')
            if c.execute('SELECT 1 FROM import_log WHERE source_hash=?',(digest,)).fetchone():
                return {'already_imported':True,'counts':{k:len(v) for k,v in bundle.items()},'archive':0}
            settings=json.loads(c.execute('SELECT data FROM settings WHERE id=1').fetchone()[0])
            if 'الإعدادات والقوائم' in wb.sheetnames:
                s=wb['الإعدادات والقوائم']
                if isinstance(s['A2'].value,(int,float)) and isinstance(s['A3'].value,(int,float)) and isinstance(s['A4'].value,(int,float)):
                    settings['thresholds']=[int(s.cell(r,1).value) for r in [2,3,4]]
                    settings['labels']=[code(s.cell(r,2).value) for r in [2,3,4,5]]
                    settings['reasons']=[code(s.cell(r,3).value) for r in range(2,min(s.max_row,100)+1) if s.cell(r,3).value]
                    counter=s['E2'].value
                else:
                    counter=s['C2'].value if isinstance(s['C2'].value,(int,float)) else s['D1'].value
                if isinstance(counter,(int,float)) and counter>=1:settings['next_serial']=max(settings['next_serial'],int(counter))
            if sorted(settings['thresholds'])!=settings['thresholds'] or len(set(settings['thresholds']))!=3 or not all(settings['labels']):
                raise ValueError('حدود أو مسميات التصنيف في المصنف غير صالحة.')
            store._replace_masters(c,bundle)
            added=0
            if 'Archive' in wb.sheetnames:
                s=wb['Archive']
                for values in s.iter_rows(min_row=4,max_col=15,values_only=True):
                    if not isinstance(values[0],(int,float)) and not code(values[0]).isdigit():continue
                    serial=int(values[0]);existing=c.execute('SELECT data FROM receipts WHERE serial=?',(serial,)).fetchone()
                    if existing:
                        prior=json.loads(existing['data'])
                        if prior.get('customer',{}).get('code')!=code(values[2]) or code(prior.get('return_number',''))!=code(values[10]):
                            raise Conflict(f'تعارض السند {serial} مع الأرشيف الموجود؛ لم يُستبدل أي سجل.')
                        continue
                    created=values[1].isoformat() if isinstance(values[1],datetime) else (iso(values[1]) or now().isoformat())
                    totals={k:int(values[i] or 0) for k,i in [('cartons',5),('boxes',6),('pieces',7)]}
                    snapshot={'serial':serial,'created_at':created,'customer':{'code':code(values[2]),'name':code(values[3])},'driver':{'code':code(values[12]),'name':code(values[13])},'return_number':code(values[10]),'invoice_date':iso(values[11]),'operator':'استيراد Excel','items':[],'totals':totals,'settings':settings.copy(),'legacy_count':int(values[4] or 0),'legacy_pdf_name':code(values[9]),'legacy_warehouses':code(values[14])}
                    # Preserve detailed RET_ worksheets when present; never invent missing item data.
                    for sheetname in wb.sheetnames:
                        if sheetname.startswith('RET_') and sheetname[4:].isdigit() and int(sheetname[4:])==serial:
                            saved=wb[sheetname]
                            for row in saved.iter_rows(min_row=11,max_row=min(saved.max_row,1000),max_col=13,values_only=True):
                                if row[1] and row[2] and iso(row[9]):
                                    snapshot['items'].append({'product_code':code(row[1]),'name':code(row[2]),'cartons':int(row[6] or 0),'boxes':int(row[7] or 0),'pieces':int(row[8] or 0),'expiry':iso(row[9]),'status':code(row[10]),'status_index':settings['labels'].index(code(row[10])) if code(row[10]) in settings['labels'] else 3,'reason':code(row[11]),'notes':code(row[12])})
                    c.execute('INSERT INTO receipts(serial,request_id,hash,created_at,data,pdf,print_state,legacy) VALUES(?,?,?,?,?,NULL,?,1)',(serial,'import:'+digest+':'+str(serial),digest,created,dump(snapshot),code(values[8]) or 'مستورد'))
                    added+=1
            maximum=c.execute('SELECT coalesce(max(serial),0) FROM receipts').fetchone()[0]
            settings['next_serial']=max(settings['next_serial'],maximum+1)
            c.execute('UPDATE settings SET data=? WHERE id=1',(dump(settings),))
            c.execute('INSERT INTO import_log VALUES(?,?)',(digest,now().isoformat()))
            store.audit(c,'workbook_import',{'hash':digest,'counts':{k:len(v) for k,v in bundle.items()},'archive':added})
        return {'already_imported':False,'counts':{k:len(v) for k,v in bundle.items()},'archive':added}
    finally:wb.close()
