"""Validate Excel's macro-project linkage, including the namespace regression."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E
import argparse

VBA_TYPE='http://schemas.microsoft.com/office/2006/relationships/vbaProject'
BOOK_MIME='application/vnd.ms-excel.sheet.macroEnabled.main+xml'

def check(path):
    with ZipFile(path) as z:
        assert z.testzip() is None,'Invalid XLSM ZIP'
        relationships=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        vba=[r for r in relationships if r.get('Type')==VBA_TYPE]
        assert len(vba)==1,'Missing or incorrect Microsoft VBA relationship namespace'
        target=vba[0].get('Target')
        part=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        assert part in z.namelist(),'VBA target not present'
        assert z.read(part).startswith(bytes.fromhex('D0CF11E0A1B11AE1')),'Invalid VBA binary'
        types=E.fromstring(z.read('[Content_Types].xml'))
        assert any(t.get('PartName')=='/xl/workbook.xml' and t.get('ContentType')==BOOK_MIME for t in types),'Workbook must be macro-enabled'
        assert any(t.get('PartName')=='/'+part and t.get('ContentType')=='application/vnd.ms-office.vbaProject' for t in types),'VBA content type missing'
    return True

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('file',nargs='?',type=Path,default=Path(__file__).resolve().parents[1]/'Product_Calculator_Search_v2.xlsm')
    args=parser.parse_args();check(args.file);print('PASS: macro-enabled workbook with a valid Microsoft VBA project relationship.')
