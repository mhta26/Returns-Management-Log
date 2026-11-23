"""Regression checks for calculator VBA call signatures and duplicated UI cases.
Requires oletools. These source checks are not an Excel/VBE compiler.
"""
from pathlib import Path
import argparse,re
from oletools.olevba import VBA_Parser


def args(text):
    depth=0;quoted=False;i=0;count=1 if text.strip() else 0
    while i<len(text):
        char=text[i]
        if char=='"':
            if quoted and i+1<len(text) and text[i+1]=='"':i+=2;continue
            quoted=not quoted
        elif not quoted:
            if char=='(':depth+=1
            elif char==')':depth-=1
            elif char==',' and depth==0:count+=1
        i+=1
    return count


def parentheses(text,start):
    depth=1;quoted=False;i=start+1
    while i<len(text):
        char=text[i]
        if char=='"':
            if quoted and i+1<len(text) and text[i+1]=='"':i+=2;continue
            quoted=not quoted
        elif not quoted:
            if char=='(':depth+=1
            elif char==')':
                depth-=1
                if depth==0:return text[start+1:i]
        i+=1
    raise AssertionError('Unclosed VBA call parentheses')


def issues(sources):
    module=sources['Module1.bas'].replace('\r\n','\n');form=sources['SearchWindow.frm'].replace('\r\n','\n');errors=[]
    signature=re.search(r'Public Sub Configure\(([^\n]+)\)',form)
    expected=args(signature.group(1))
    calls=re.findall(r'^\s*dialog\.Configure (.+)$',module,re.M)
    if len(calls)!=1:errors.append('Expected exactly one dialog.Configure call')
    for call in calls:
        actual=args(call)
        if actual!=expected:errors.append(f'dialog.Configure receives {actual} arguments but expects {expected}')
    block=re.search(r'Public Function Caption\b.*?End Function',module,re.S).group()
    cases=re.findall(r'^\s*Case (\d+):',block,re.M)
    if len(cases)!=len(set(cases)):errors.append('Duplicate Caption Case identifiers')
    ranges={'AddLabel':(6,8),'Caption':(1,1),'Display':(1,2),'Normalize':(1,1),'U':(1,1)}
    for name,(low,high) in ranges.items():
        for source in [module,form]:
            for match in re.finditer(r'\b'+name+r'\(',source):
                prefix=source[source.rfind('\n',0,match.start())+1:match.start()]
                if re.search(r'\b(?:Function|Sub)\s*$',prefix):continue
                actual=args(parentheses(source,match.end()-1))
                if not low<=actual<=high:errors.append(f'{name} receives {actual} arguments, expected {low}..{high}')
    return errors


def check(path):
    sources={n:c for _,_,n,c in VBA_Parser(str(path)).extract_macros()}
    found=issues(sources)
    assert not found,'; '.join(found)
    return True

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('file',type=Path);params=parser.parse_args();check(params.file)
    print('PASS: Configure call matches form signature; no duplicate UI cases; helper call arities are consistent.')
