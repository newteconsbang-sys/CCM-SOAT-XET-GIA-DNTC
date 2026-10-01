import sys,unittest,json,uuid
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'backend/vendor'));sys.path.insert(0,str(ROOT))
from fastapi.testclient import TestClient
from backend.app import app,TOKEN
from openpyxl import load_workbook
class ServiceTests(unittest.TestCase):
 def setUp(self):self.c=TestClient(app,base_url='https://localhost:3443',headers={'X-CCM-Token':TOKEN})
 def test_security(self):
  self.assertEqual(self.c.get('/api/health').json()['ai'],'disabled')
  self.assertEqual(TestClient(app,base_url='https://localhost:3443').get('/api/prices').status_code,401)
  self.assertEqual(self.c.post('/api/profiles',headers={'origin':'https://evil.example'},json={}).status_code,403)
 def test_rule_versions(self):
  r=self.c.get('/api/rules',params={'business':'NTP'}).json();r['effective']='2026-10-01'
  saved=self.c.post('/api/rules',json=r).json();self.assertEqual(saved['version'],r['version']+1)
  r['highPriceIsError']=True;self.assertEqual(self.c.post('/api/rules',json=r).status_code,400)
 def test_backup_durable(self):
  bid=str(uuid.uuid4());p={'id':bid,'workbook':'TEST-WORKBOOK','before':[{'sheet':'A','address':'A1','values':[[100]],'formulas':[['=10*10']]}]}
  self.assertEqual(self.c.post('/api/backups',json=p).status_code,200)
  self.assertEqual(self.c.get('/api/backups/'+bid).json()['before'][0]['values'],[[100]])
  self.assertEqual(self.c.post('/api/backups',json=p).status_code,409)
 def test_export(self):
  headers=['STT','Hạng mục','ĐVT','KL','Chào VT','Chào NC','Soát VT','Soát NC','Chênh','%','TT soát','Ảnh hưởng','Trạng thái','Lý do','Nguồn']
  row=['1.10','Ống đồng Ø20','m',10,100,20,90,18,12,12/108*100,1080,120,'Cần kiểm tra','Cùng phạm vi','Nguồn minh họa']
  for fmt in ['xlsx','docx','pdf']:
   response=self.c.post('/api/reports',json={'business':'NCC','title':'BÁO CÁO KIỂM THỬ – DỮ LIỆU MINH HỌA','project':'KIỂM THỬ','format':fmt,'columns':list(range(15)),'sections':['Cơ sở','Kết quả'],'headers':headers,'rows':[row]})
   self.assertEqual(response.status_code,200,response.text);name=response.json()['filename'];path=ROOT/'reports'/name;self.assertTrue(path.stat().st_size>100)
   if fmt=='xlsx':
    w=load_workbook(path);s=w.active;self.assertEqual(s.cell(s.max_row,1).value,'1.10');self.assertEqual(s.cell(s.max_row,11).value,1080);self.assertEqual(s.cell(s.max_row,12).value,120)
 def test_reference_validation(self):self.assertEqual(self.c.post('/api/prices',json={'entries':[{'name':'A','unit':'m','price':10}]}).status_code,400)
if __name__=='__main__':unittest.main(verbosity=2)
