from __future__ import annotations
import json
from pathlib import Path
from urllib.parse import urlencode
import httpx
from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal, Slot, Qt


class Api:
    def __init__(self,url,token=''):
        self.url=url.rstrip('/')
        if not self.url.startswith(('http://','https://')):raise ValueError('أدخل رابط الخادم مثل http://192.168.1.10:8765')
        self.token=token
        self.http=httpx.Client(timeout=httpx.Timeout(60,connect=6),trust_env=False)

    def call(self,method,path,data=None,content=None,binary=False):
        headers={'Authorization':'Bearer '+self.token} if self.token else {}
        if content is not None:headers['Content-Type']='application/octet-stream'
        try:r=self.http.request(method,self.url+path,json=data,content=content,headers=headers)
        except httpx.RequestError as e:raise RuntimeError('تعذر الاتصال بالخادم. تحقق من عنوانه وتشغيله واتصال الشبكة؛ لم تُحذف مدخلاتك.') from e
        if not r.is_success:
            try:detail=r.json().get('detail',r.text)
            except ValueError:detail=r.text
            if isinstance(detail,list):detail='؛ '.join(str(x.get('msg','بيانات غير صالحة')) for x in detail)
            raise RuntimeError(str(detail))
        return r.content if binary else r.json()


class Job(QRunnable):
    def __init__(self,runner,key,fn):super().__init__();self.runner=runner;self.key=key;self.fn=fn
    def run(self):
        try:self.runner.completed.emit((self.key,self.fn(),None))
        except Exception as e:self.runner.completed.emit((self.key,None,str(e)))


class Tasks(QObject):
    completed=Signal(object)
    def __init__(self,parent=None):
        super().__init__(parent);self.pool=QThreadPool(self);self.callbacks={};self.key=0
        self.completed.connect(self.done,Qt.ConnectionType.QueuedConnection)
    def submit(self,fn,ok,error=None):
        self.key+=1;self.callbacks[self.key]=(ok,error);self.pool.start(Job(self,self.key,fn))
    @Slot(object)
    def done(self,result):
        key,value,error=result;ok,fail=self.callbacks.pop(key)
        if error:
            if fail:fail(error)
        else:ok(value)
