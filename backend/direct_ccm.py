from fastapi import APIRouter,HTTPException
import math
import json,re
import uuid
from pathlib import Path
from backend.price_expressions import expression_fields, expression_conditions
from backend.library_bridge import snapshot

def make_router(records,insert,business,library_root=None):
    router=APIRouter(prefix='/api/ccm')
    root=Path(library_root) if library_root is not None else Path(__file__).resolve().parents[1]/'libraries'
    def folder(b):return root/('DNTC' if b=='Điện nước ĐNTC' else b)
    @router.get('/library')
    def library(business:str):
        b=globals_business(business)
        f=folder(b)
        if (f/'active.json').exists():
            version=json.loads((f/'active.json').read_text(encoding='utf-8'))['version']
            if not re.fullmatch(r'[A-Za-z0-9-]+',str(version)):raise HTTPException(400,'Phiên bản thư viện không hợp lệ.')
            document=json.loads((f/'versions'/f'{version}.json').read_text(encoding='utf-8'))
            return dict(document,id=str(version),entries=[dict(p,brand=p.get('brand',''),material=p.get('material'),labor=p.get('labor')) for p in document['entries']])
        rows=records('ccm-library',b)
        return rows[-1] if rows else {'entries':[],'version':0,'id':None}
    globals_business=business
    @router.get('/library-snapshot')
    def library_snapshot(business:str):
        b=globals_business(business)
        try:
            return snapshot(library(b),b)
        except (OSError,KeyError,ValueError,TypeError):
            raise HTTPException(409,'Không đọc được thư viện CCM-AI; kiểm tra file hoặc nạp lại thư viện.')
    @router.post('/library')
    def save(body:dict):
        b=business(body.get('business'));entries=body.get('entries')
        if body.get('confirmed') is not True or not isinstance(entries,list) or not 1<=len(entries)<=5000:
            raise HTTPException(400,'Xác nhận ánh xạ thư viện; tối đa 5.000 dòng.')
        for p in entries:
            if not isinstance(p,dict):raise HTTPException(400,'Dòng thư viện không hợp lệ.')
            try:expressions=expression_fields(p)
            except ValueError as error:raise HTTPException(400,str(error))
            if not all(isinstance(p.get(k),str) and p[k].strip() for k in ['description','source']) or not isinstance(p.get('unit'),str) or not p['unit'].strip() and not expressions:
                raise HTTPException(400,'Mỗi dòng cần mô tả, ĐVT và nguồn.')
            prices=[p.get(k) for k in ['material','labor']]
            quantity=p.get('quantity')
            if quantity is not None and (type(quantity) not in [int,float] or not math.isfinite(quantity) or quantity<0):
                raise HTTPException(400,'Khối lượng phải là số hữu hạn không âm hoặc để trống.')
            if all(v is None for v in prices) and not expressions or any(v is not None and (type(v) not in [int,float] or not math.isfinite(v) or v<0) for v in prices):
                raise HTTPException(400,'Đơn giá phải là số hữu hạn không âm; giá thiếu để trống.')
            p.update(expressions)
            if expressions:p['conditions']=expression_conditions(p)
        saved=insert('ccm-library',b,{'business':b,'entries':entries,'name':str(body.get('name','Thư viện giá'))[:150]})
        f=folder(b);(f/'versions').mkdir(parents=True,exist_ok=True)
        version=f'app-{saved["id"]}-v{saved["version"]}'
        saved=dict(saved,id=version,version=version)
        (f/'versions'/f'{version}.json').write_text(json.dumps(saved,ensure_ascii=False,indent=2),encoding='utf-8')
        temporary=f/('active-'+uuid.uuid4().hex+'.tmp');temporary.write_text(json.dumps({'version':version}),encoding='utf-8');temporary.replace(f/'active.json')
        return saved
    @router.get('/settings')
    def settings(business:str):
        r=records('ccm-settings',globals_business(business));return r[-1] if r else {}
    @router.post('/settings')
    def save_settings(body:dict):
        columns=body.get('columns',{})
        if set(columns)!={'description','brand','unit','material','labor'} or any(not isinstance(v,str) or not re.fullmatch('[A-Za-z]{1,3}',v) for v in columns.values()) or len(set(v.upper() for v in columns.values()))!=5:
            raise HTTPException(400,'Chọn năm cột khác nhau cho nguồn và giá.')
        return insert('ccm-settings',business(body.get('business')),body)
    from backend.web_intake import make_router as intake_router
    router.include_router(intake_router(root.parent,globals_business,save))
    @router.post('/skill-call')
    def skill_call(body:dict):
        from backend.price_service import LibraryService
        try:
            return LibraryService(root).call(body.get('name'),body.get('arguments',{}))
        except (ValueError,KeyError,TypeError,OSError) as error:
            raise HTTPException(400,str(error))
    return router
