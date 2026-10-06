from html.parser import HTMLParser
from pathlib import Path

class ReportParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.ids=[]; self.links=[]; self.bank=False; self.in_body=False
        self.cells=[]; self.bank_rows=[]; self.cases=0
    def handle_starttag(self,tag,attrs):
        a=dict(attrs)
        if 'id' in a: self.ids.append(a['id'])
        if tag=='a' and a.get('href','').startswith('#'): self.links.append(a['href'][1:])
        if tag=='section' and a.get('class')=='case': self.cases+=1
        if tag=='table' and a.get('class')=='bank': self.bank=True
        if self.bank and tag=='tbody': self.in_body=True
        if self.bank and self.in_body and tag=='tr': self.cells=[]
        if self.bank and self.in_body and tag=='td': self.cells.append(a.get('class',''))
    def handle_endtag(self,tag):
        if self.bank and self.in_body and tag=='tr': self.bank_rows.append(self.cells)
        if self.bank and tag=='tbody': self.in_body=False
        if self.bank and tag=='table': self.bank=False

text=Path('Результат проверки реальных файлов/Отчёт.html').read_text(encoding='utf-8-sig')
p=ReportParser(); p.feed(text)
assert len(p.ids)==len(set(p.ids))
assert all(a in p.ids for a in p.links)
assert len(p.bank_rows)==18 and p.cases==18
assert all(len(row)==6 and all(c in ('nowrap','money') for c in row[:4]) for row in p.bank_rows)
assert 'white-space:nowrap!important' in text and 'overflow-x:auto' in text
print('PASS: 18 case sections, all local links resolve, first four columns protected from wrapping')
