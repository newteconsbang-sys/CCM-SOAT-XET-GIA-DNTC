"""Opt-in real Ollama test; synthetic data only, stores a local result in SQLite."""
import json,ssl,urllib.request
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
INSTALL=Path(r'D:\00.Setupgpt\SoatXetGia')
context=ssl.create_default_context(cadata=ssl.DER_cert_to_PEM_cert((INSTALL/'config/localhost.cer').read_bytes()))
token=''
def api(path,body=None):
    req=urllib.request.Request('https://localhost:3443/api/'+path,
        data=None if body is None else json.dumps(body,ensure_ascii=False).encode(),
        headers={'X-CCM-Token':token,'Content-Type':'application/json'})
    try:
        with urllib.request.urlopen(req,context=context,timeout=360) as response:return json.load(response)
    except urllib.error.HTTPError as e:
        raise RuntimeError(e.read().decode()) from e
if __name__=='__main__':
    token=api('session')['token']
    q={'id':'1.10','description':'Ống nhựa uPVC Ø20 PN10, cùng hãng ABC, cung cấp vật tư, chưa VAT','unit':'m',
       'quantity':10,'material':22000,'labor':0,'row':5,'refs':{'material':"'TEST Quote'!E6"}}
    r={**q,'material':20000,'row':7,'refs':{'material':"'TEST Reference'!E8",'labor':"'TEST Reference'!F8"}}
    p=api('ai/prepare',{'business':'NCC','quote':q,'references':[r],'priceIds':[],'documentIds':[]})
    a=api('ai/analyze',{'planId':p['planId'],'digest':p['digest'],'confirmed':True})
    (ROOT/'docs/local-ai-live-review.json').write_text(json.dumps(a,ensure_ascii=False,indent=2),encoding='utf-8')
    assert a['reference'] and a['reference']['material']==20000, 'AI did not propose the consistent supplied reference'
    assert a['status']=='yellow' and a['materialPrice'] is None
    api('ai/feedback',{'resultId':a['resultId'],'decision':'rejected','note':'Dữ liệu kiểm thử minh họa; không áp dụng workbook.'})
    print(json.dumps(a,ensure_ascii=True))
