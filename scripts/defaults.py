from pathlib import Path
import json,sqlite3
ROOT=Path(__file__).resolve().parents[1]
con=sqlite3.connect(ROOT/'data/ccm.sqlite3')
for kind,name,payload in con.execute("SELECT kind,name,payload FROM records WHERE kind IN ('rule','template') AND version=1"):
 p=json.loads(payload);p.pop('updated',None)
 slug={'Điện nước ĐNTC':'dntc','EW':'ew','NTP':'ntp','NCC':'ncc'}[name]
 folder=ROOT/('rules' if kind=='rule' else 'templates')
 (folder/f'default-{slug}.json').write_text(json.dumps(p,ensure_ascii=False,indent=2),encoding='utf-8')
con.close()
