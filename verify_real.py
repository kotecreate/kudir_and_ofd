"""Independent read-only reconciliation of the supplied layouts; no workbook edits."""
import json,re,hashlib
from pathlib import Path
from collections import defaultdict,Counter
from decimal import Decimal as D
from datetime import datetime
import openpyxl

root=Path(__file__).parent
kpath=next(root.glob('*_2026_1_*.xlsx')); opath=next(root.glob('*2026-10-01*.xlsx'))
money=lambda v:D(str(v or 0)).quantize(D('.01'))
kbook=openpyxl.load_workbook(kpath,data_only=True,read_only=True)
obook=openpyxl.load_workbook(opath,data_only=True,read_only=True)
k=[]; o=[]; cash=defaultdict(lambda:D(0)); bank=defaultdict(lambda:D(0)); weird=[]
for s in kbook:
    if not s.title.startswith('Квартал-'): continue
    for n,row in enumerate(s.values,1):
        if not isinstance(row[0],(int,float)) or not isinstance(row[1],str) or not re.match(r'\d{2}\.\d{2}\.\d{4}',row[1]): continue
        doc=datetime.strptime(row[1][:10],'%d.%m.%Y').date().isoformat()
        desc=row[2] or ''; a=money(row[3]); sale=doc; kind='unknown'
        if desc.startswith('Оплата наличными'):
            kind='cash'; cash[doc[:7]]+=a
        elif 'по терминалу' in desc and (m:=re.search(r' за (\d{4}-\d{2}-\d{2})',desc)):
            kind='card'; sale=m[1]; bank[sale]+=a
        elif desc.startswith('Зачисление по QR коду'):
            kind='qr'; bank[sale]+=a
        else: weird.append((s.title,n,desc,a))
        k.append(dict(sheet=s.title,row=n,doc=doc,date=sale,kind=kind,amount=a))
for n,row in enumerate(obook.worksheets[0].values,1):
    if len(row)>19 and row[13]=='Кассовый чек':
        o.append(dict(row=n,date=row[10].date().isoformat(),kind=row[16],amount=money(row[17]),cash=money(row[18]),bank=money(row[19])))
start=min(r['date'] for r in o); end=max(r['date'] for r in o)
ob=defaultdict(lambda:D(0)); oc=defaultdict(lambda:D(0))
for r in o:
    if r['kind'] not in ('Приход','Возврат прихода'): continue
    sign=-1 if r['kind']=='Возврат прихода' else 1
    ob[r['date']]+=r['bank']*sign; oc[r['date'][:7]]+=r['cash']*sign
diffs=[dict(date=d,k=bank[d],o=ob[d],diff=ob[d]-bank[d]) for d in sorted(bank.keys()|ob.keys()) if bank[d]!=ob[d]]
cashdiff=[dict(month=m,k=cash[m],o=oc[m],diff=oc[m]-cash[m]) for m in sorted(cash.keys()|oc.keys())]
result=dict(kfile=str(kpath),ofile=str(opath),hashes={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in (kpath,opath)},start=start,end=end,kcount=len(k),ocount=len(o),kall=sum(r['amount'] for r in k),kledger=sum(r['amount'] for r in k if start<=r['doc']<=end),kaligned=sum(r['amount'] for r in k if start<=r['date']<=end),o_receipts=sum(r['amount'] for r in o if r['kind']=='Приход'),o_other=sum(r['amount'] for r in o if r['kind'] not in ('Приход','Возврат прихода')),cash=cashdiff,bank_diffs=diffs,unknown=weird,k=k,o=o)
out=root/'test-output'/'independent-real.json'; out.write_text(json.dumps(result,default=str,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({key:value for key,value in result.items() if key not in ('k','o','hashes')},default=str,ensure_ascii=False,indent=2))
