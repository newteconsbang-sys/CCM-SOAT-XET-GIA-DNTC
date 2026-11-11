import sys, unittest, tempfile, io, json, subprocess, shutil
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'backend/vendor'));sys.path.insert(0,str(ROOT))
from fastapi import FastAPI,HTTPException
from fastapi.testclient import TestClient
from openpyxl import Workbook
from backend.direct_ccm import make_router
from backend.price_service import LibraryService

class WebTests(unittest.TestCase):
 def setUp(self):
  self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name);self.serial=0;self.db=[]
  def business(value):
   if value not in ['Điện nước ĐNTC','EW','NCC','NTP']:raise HTTPException(400,'Sai nghiệp vụ')
   return value
  def insert(kind,b,p):
   self.serial+=1;item=dict(p,id=self.serial,version=self.serial);self.db.append((kind,b,item));return item
  def records(kind,b):return [p for k,selected,p in self.db if k==kind and selected==b]
  app=FastAPI();app.include_router(make_router(records,insert,business,self.root/'libraries'));self.client=TestClient(app)
 def tearDown(self):self.temp.cleanup()
 def upload(self,raw=None,name='prices.csv',business='NTP'):
  if raw is None:raw='Mô tả,ĐVT,Hãng,Vật tư,Nhân công\nỐng DN25,m,ABC,0,18000\nThi công trọn gói,Lô,,0,9000\nChỉ có nhân công,m,,,12000\nNhóm,,,,\n'.encode('utf-8')
  response=self.client.post('/api/ccm/intake/upload',files={'file':(name,raw)},data={'selected_business':business})
  self.assertEqual(response.status_code,200,response.text);return response.json()
 def body(self,d):return dict(id=d['id'],business=d['business'],sheet=d['sheets'][0]['name'],columns={'description':0,'unit':1,'brand':2,'material':3,'labor':4},headerRow=1,lastRow=5,name='Test prices')
 def test_upload_preview_commit_skill_reads_only_selected_library(self):
  d=self.upload();self.assertEqual(self.db,[])
  p=self.client.post('/api/ccm/intake/preview',json=self.body(d));self.assertEqual(p.status_code,200,p.text)
  self.assertEqual(p.json()['count'],3);self.assertEqual(p.json()['skippedRows'],[5])
  self.assertEqual(p.json()['entries'][0]['material'],0)
  self.assertIsNone(p.json()['entries'][2]['material'])
  saved=self.client.post('/api/ccm/intake/commit',json=dict(self.body(d),sourceHash=d['sha256'],confirmed=True))
  self.assertEqual(saved.status_code,200,saved.text)
  service=LibraryService(self.root/'libraries');read=service.call('ccm_read_library',dict(business='NTP'))
  self.assertEqual(read['version'],saved.json()['version']);self.assertEqual(read['total'],3)
  self.assertEqual(read['entries'][0]['labor'],18000);self.assertIn(d['sha256'],read['entries'][0]['source'])
  results=service.call('ccm_review_rows',dict(business='NTP',version=read['version'],rows=[dict(row=12,description='Ống DN25',unit='m',brand='ABC'),dict(row=13,description='Thi công trọn gói',unit='Lô',brand=''),dict(row=14,description='Ống DN32',unit='m',brand='ABC')]))
  self.assertEqual([p['status'] for p in results['results']],['matched','lot_deferred','unmatched'])
  self.assertEqual(service.call('ccm_list_libraries',{})['libraries'][2]['status'],'not_loaded')
  self.assertFalse((self.root/'libraries/NCC/active.json').exists())
  (self.root/'backend').mkdir();(self.root/'backend/__init__.py').write_text('')
  shutil.copyfile(ROOT/'backend/price_service.py',self.root/'backend/price_service.py')
  helper=ROOT/'skills/soatxetgia-ccm/scripts/read_web.py'
  cli=subprocess.run([sys.executable,str(helper),'--app-root',str(self.root),'--action','read','--business','NTP'],capture_output=True,text=True,encoding='utf-8',timeout=10)
  self.assertEqual(cli.returncode,0,cli.stderr);self.assertEqual(json.loads(cli.stdout)['entries'][0]['material'],0)
  again=self.client.post('/api/ccm/intake/commit',json=dict(self.body(d),sourceHash=d['sha256'],confirmed=True))
  self.assertEqual(again.status_code,200)
  self.assertNotEqual(again.json()['version'],read['version'])
  with self.assertRaises(ValueError):service.call('ccm_review_rows',dict(business='NTP',version=read['version'],rows=[dict(row=12,description='Ống DN25',unit='m',brand='ABC')]))
 def test_wrong_business_and_changed_source_are_blocked(self):
  d=self.upload();b=self.body(d)
  self.assertEqual(self.client.post('/api/ccm/intake/preview',json=dict(b,business='NCC')).status_code,400)
  path=self.root/'uploads/price-library'/d['id']/'original.csv';path.write_bytes(b'changed')
  self.assertEqual(self.client.post('/api/ccm/intake/commit',json=dict(b,sourceHash=d['sha256'],confirmed=True)).status_code,409)
 def test_formula_in_price_is_not_executed_and_mapping_is_required(self):
  d=self.upload(b'Mo ta,DVT,Hang,VT,NC\nPipe,m,,=2*2,10\n',name='formula.csv');b=dict(self.body(d),lastRow=2)
  self.assertEqual(self.client.post('/api/ccm/intake/preview',json=b).status_code,400)
  self.assertEqual(self.client.post('/api/ccm/intake/commit',json=dict(b,sourceHash=d['sha256'])).status_code,400)
 def test_xlsx_selects_header_and_typed_prices(self):
  book=Workbook();sheet=book.active;sheet.title='Thư viện';sheet.append(['Tên công trình']);sheet.append(['Mô tả','ĐVT','Hãng','VT','NC']);sheet.append(['Cáp 2x2.5','m','ABC',32000,None]);out=io.BytesIO();book.save(out)
  d=self.upload(out.getvalue(),'test.xlsx');b=dict(self.body(d),headerRow=2,lastRow=3)
  self.assertEqual(self.client.get(f'/api/ccm/intake/{d["id"]}/sheet',params=dict(sheet='Thư viện',header_row=2)).json()['headers'][0],'Mô tả')
  p=self.client.post('/api/ccm/intake/preview',json=b);self.assertEqual(p.status_code,200,p.text);self.assertEqual(p.json()['entries'][0]['material'],32000)
 def test_json_business_mismatch_and_input_path_are_rejected(self):
  response=self.client.post('/api/ccm/intake/upload',files={'file':('wrong.json',json.dumps({'business':'NCC','entries':[{'description':'Pipe','unit':'m','material':1}]}))},data={'selected_business':'EW'})
  self.assertEqual(response.status_code,400)
  self.assertEqual(self.client.get('/api/ccm/intake/../sheet',params={'sheet':'CSV'}).status_code,404)

if __name__=='__main__':unittest.main()
