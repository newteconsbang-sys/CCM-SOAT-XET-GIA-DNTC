"""Transport is mocked here; real inference has a separate installed-machine receipt."""
import sys,unittest,tempfile,json,sqlite3
from pathlib import Path
from contextlib import closing
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'backend/vendor'));sys.path.insert(0,str(ROOT))
from fastapi import FastAPI
from fastapi.testclient import TestClient
from backend.local_ai import make_router

class LocalAITests(unittest.TestCase):
 def setUp(self):
  self.tmp=tempfile.TemporaryDirectory(dir=ROOT/"data");self.root=Path(self.tmp.name);(self.root/'config').mkdir()
  (self.root/'config/ai.json').write_text(json.dumps({'enabled':True,'model':'qwen3:4b'}))
  self.saved=[];self.answer={'referenceKey':None,'materialSource':None,'laborSource':None,'explanation':'Chưa đủ dữ liệu','warnings':[],'evidenceIds':[]}
  self.data={'price':[{'id':9,'version':1,'name':'Ống PVC 20','unit':'m','price':100,'source':'TEST ONLY'}],'rule':[{'version':1,'checks':['VAT']}],'document':[]}
  def records(kind,name=None): return self.data.get(kind,[])
  def insert(kind,name,payload):
   payload={**payload,'id':100+len(self.saved),'version':1+len(self.saved),'updated':'2026-10-01'};self.saved.append((kind,payload));self.data.setdefault(kind,[]).append(payload);return payload
  app=FastAPI();app.include_router(make_router(self.root,records,insert,lambda b:b));self.client=TestClient(app)
  def transport(path,body=None,timeout=10):
   if path=='tags':return {'models':[{'name':'qwen3:4b'}]}
   if path=='show':return {}
   return {'done':True,'message':{'content':json.dumps(self.answer)}}
  self.mock=patch('backend.local_ai.ollama',side_effect=transport);self.mock.start();self.addCleanup(self.mock.stop);self.addCleanup(self.tmp.cleanup)
 def prepare(self):return self.client.post('/api/ai/prepare',json={'business':'NCC','quote':{'description':'Ống PVC 20','unit':'m'},'references':[],'priceIds':[9]}).json()
 def analyze(self,p,confirmed=True):return self.client.post('/api/ai/analyze',json={'planId':p['planId'],'digest':p['digest'],'confirmed':confirmed})
 def test_consent_and_replay(self):
  p=self.prepare();self.assertEqual(self.analyze(p,False).status_code,400)
  self.assertEqual(self.analyze(p).status_code,200);self.assertEqual(self.analyze(p).status_code,409)
 def test_price_source_is_authoritative(self):
  self.answer.update(materialSource='price:9',evidenceIds=['price:9'])
  result=self.analyze(self.prepare()).json();self.assertEqual(result['materialPrice']['price'],100);self.assertEqual(result['status'],'yellow')
  self.assertEqual(self.client.post('/api/ai/feedback',json={'resultId':result['resultId'],'decision':'rejected'}).json()['automaticTraining'],False)
 def test_invented_sources_are_blocked(self):
  self.answer['materialSource']='price:999';self.assertEqual(self.analyze(self.prepare()).status_code,502);self.assertEqual(self.saved,[])
 def test_model_cannot_generate_price(self):
  self.answer['price']=123;self.assertEqual(self.analyze(self.prepare()).status_code,502)
 def test_cloud_setting_is_blocked(self):
  self.assertEqual(self.client.post('/api/ai/settings',json={'model':'qwen3:cloud'}).status_code,400)
 def test_remote_model_metadata_is_blocked(self):
  def remote(path,body=None,timeout=10):return {'models':[{'name':'qwen3:4b'}]} if path=='tags' else {'remote_host':'https://example.test'}
  with patch('backend.local_ai.ollama',side_effect=remote):
   self.assertEqual(self.client.post('/api/ai/settings',json={'model':'qwen3:4b'}).status_code,400)
 def test_setting_changed_invalidates_preview(self):
  p=self.prepare();(self.root/'config/ai.json').write_text(json.dumps({'enabled':False,'model':'qwen3:4b'}));self.assertEqual(self.analyze(p).status_code,409)
 def test_test_uses_synthetic_data(self):
  r=self.client.post('/api/ai/test',json={'model':'qwen3:4b'});self.assertTrue(r.json()['synthetic'])
 def memory(self,business='NCC',unit='m',approved=True):
  rate={'mode':'single','value':100,'source':'TEST SOURCE'}
  return self.client.post('/api/ai/memories',json={'business':business,'approved':approved,'lesson':'Kiểm tra phụ kiện đã bao gồm trong giá ống hay chưa.',
   'review':{'confirmed':True,'quote':{'description':'Ống PVC 20','unit':unit,'errors':[]},'material':rate,'labor':rate,'reason':'Cùng thông số nhưng cần so sánh phụ kiện','status':'yellow'}})
 def test_memory_requires_explicit_approval(self):
  self.assertEqual(self.memory(approved=False).status_code,400);self.assertEqual(self.saved,[])
 def test_memory_retrieval_omits_historical_prices(self):
  self.assertEqual(self.memory().status_code,200)
  p=self.prepare();e=p['context']['approvedExamples'];self.assertEqual(len(e),1);self.assertNotIn('price',json.dumps(e));self.assertNotIn('value',json.dumps(e))
  self.answer['evidenceIds']=[e[0]['key']];self.assertEqual(self.analyze(p).status_code,200)
 def test_memory_business_unit_and_opt_out(self):
  self.memory('EW');self.memory(unit='cái');self.assertEqual(self.prepare()['context']['approvedExamples'],[])
  self.memory();p=self.client.post('/api/ai/prepare',json={'business':'NCC','quote':{'description':'Ống PVC 20','unit':'m'},'references':[],'useMemories':False}).json()
  self.assertEqual(p['context']['approvedExamples'],[])
 def test_revoked_memory_invalidates_preview_and_survives_router_restart(self):
  m=self.memory().json();p=self.prepare()
  self.assertEqual(self.client.post('/api/ai/memories/revoke',json={'business':'NCC','memoryId':m['memoryId']}).status_code,200)
  self.assertEqual(self.analyze(p).status_code,409);self.assertEqual(self.prepare()['context']['approvedExamples'],[])
  app=FastAPI();app.include_router(make_router(self.root,lambda k,n=None:self.data.get(k,[]),lambda *args:None,lambda b:b))
  self.assertEqual(TestClient(app).get('/api/ai/memories?business=NCC').json(),[])
 def test_approved_memory_persists_on_disk(self):
  path=self.root/'memory.sqlite3'
  with closing(sqlite3.connect(path)) as con: con.execute('CREATE TABLE examples(id INTEGER PRIMARY KEY, name TEXT, payload TEXT)')
  def disk_records(kind,name=None):
   if kind!='ai-memory': return self.data.get(kind,[])
   with closing(sqlite3.connect(path)) as con:return [json.loads(x[0]) for x in con.execute('SELECT payload FROM examples WHERE name=?',(name,))]
  def disk_insert(kind,name,payload):
   with closing(sqlite3.connect(path)) as con:
    cur=con.execute('INSERT INTO examples(name,payload) VALUES (?,?)',(name,'{}'));p={**payload,'id':cur.lastrowid,'version':cur.lastrowid,'updated':'2026-10-01'}
    con.execute('UPDATE examples SET payload=? WHERE id=?',(json.dumps(p),p['id']));con.commit()
   return p
  def new_client():
   app=FastAPI();app.include_router(make_router(self.root,disk_records,disk_insert,lambda b:b));return TestClient(app)
  self.client=new_client();m=self.memory().json();self.client=new_client()
  self.assertEqual(self.client.get('/api/ai/memories?business=NCC').json()[0]['memoryId'],m['memoryId'])
  self.assertEqual(self.prepare()['context']['approvedExamples'][0]['memoryId'],m['memoryId'])
if __name__=='__main__':unittest.main(verbosity=2)
