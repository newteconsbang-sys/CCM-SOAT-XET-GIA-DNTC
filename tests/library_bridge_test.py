import sys, unittest, tempfile, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'backend' / 'vendor'))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from fastapi import HTTPException, FastAPI
from fastapi.testclient import TestClient
from backend.library_bridge import snapshot
from backend.direct_ccm import make_router

class BridgeTests(unittest.TestCase):
    def document(self, business='EW'):
        return dict(business=business, version='test-v1', entries=[dict(description='Ống DN25', unit='m', material=0, labor=None, source='Giá được duyệt')])

    def test_snapshot_is_stable_and_retains_missing_and_zero(self):
        d=self.document(); a=snapshot(d,'EW'); b=snapshot(d,'EW')
        self.assertEqual(a['sha256'],b['sha256'])
        self.assertEqual(a['entries'][0]['material'],0)
        self.assertIsNone(a['entries'][0]['labor'])
        d['entries'][0]['material']=12
        self.assertNotEqual(a['sha256'],snapshot(d,'EW')['sha256'])
        self.assertNotEqual(a['sha256'],snapshot(self.document('NCC'),'NCC')['sha256'])

    def test_rejects_empty_wrong_business_and_bad_prices(self):
        with self.assertRaises(HTTPException):snapshot({'entries':[]},'EW')
        with self.assertRaises(HTTPException):snapshot(self.document('NCC'),'EW')
        for value in [True, -1, float('nan'), float('inf'), '100']:
            d=self.document();d['entries'][0]['material']=value
            with self.assertRaises(HTTPException):snapshot(d,'EW')

    def test_endpoint_reads_only_selected_business(self):
        seen=[]
        def records(kind,business):
            seen.append((kind,business));return [self.document(business)]
        def business(value):
            if value not in ['EW','NCC','NTP','Điện nước ĐNTC']:raise HTTPException(400,'Nghiệp vụ sai')
            return value
        with tempfile.TemporaryDirectory() as root:
            app=FastAPI();app.include_router(make_router(records,lambda *_:None,business,root))
            response=TestClient(app).get('/api/ccm/library-snapshot?business=EW')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.json()['business'],'EW')
        self.assertEqual(seen,[('ccm-library','EW')])

    def test_addin_save_becomes_selected_snapshot_without_other_business(self):
        def business(value):return value
        def insert(kind,b,p):return dict(p,id=21,version=1)
        with tempfile.TemporaryDirectory() as root:
            app=FastAPI();app.include_router(make_router(lambda *_:[],insert,business,root))
            client=TestClient(app)
            doc=self.document('Điện nước ĐNTC')
            saved=client.post('/api/ccm/library',json=dict(doc,confirmed=True))
            self.assertEqual(saved.status_code,200)
            read=client.get('/api/ccm/library-snapshot',params={'business':'Điện nước ĐNTC'})
            self.assertEqual(read.status_code,200)
            self.assertEqual(read.json()['business'],'DNTC')
            self.assertEqual(read.json()['version'],saved.json()['version'])
            self.assertEqual(read.json()['entries'][0]['material'],0)
            self.assertEqual(client.get('/api/ccm/library-snapshot?business=NCC').status_code,409)

if __name__=='__main__':unittest.main()
