"""Local-only CCM service. Documents are evidence, never executable instructions."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib, json, secrets, sqlite3, os, math, uuid
from fastapi import FastAPI, Request, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.base import BaseHTTPMiddleware

ROOT = Path(__file__).resolve().parents[1]
BUSINESSES = ['Điện nước ĐNTC', 'EW', 'NTP', 'NCC']
TOKEN = secrets.token_urlsafe(32)
MAX_BYTES = 25 * 1024 * 1024
for name in ['data','uploads','reports','backups','rules','templates','config']:
    (ROOT/name).mkdir(exist_ok=True)
DB = ROOT/'data'/'ccm.sqlite3'
def now(): return datetime.now(timezone.utc).isoformat()
def connect():
    con=sqlite3.connect(DB)
    con.row_factory=sqlite3.Row
    return con
with connect() as con:
    con.execute('CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY, kind TEXT, name TEXT, version INTEGER, updated TEXT, payload TEXT)')
    con.execute('CREATE TABLE IF NOT EXISTS backup_index (id TEXT PRIMARY KEY, workbook TEXT, state TEXT, updated TEXT)')

def records(kind, name=None):
    with connect() as con:
        sql='SELECT * FROM records WHERE kind=?';args=[kind]
        if name is not None: sql+=' AND name=?';args.append(name)
        rows=con.execute(sql+' ORDER BY id',args).fetchall()
    return [dict(json.loads(r['payload']), id=r['id'],version=r['version'],updated=r['updated']) for r in rows]
def insert(kind,name,payload):
    with connect() as con:
        version=con.execute('SELECT COALESCE(MAX(version),0)+1 FROM records WHERE kind=? AND name=?',(kind,name)).fetchone()[0]
        stamp=now();p=dict(payload);p.pop('id',None);p['version']=version;p['updated']=stamp
        cur=con.execute('INSERT INTO records(kind,name,version,updated,payload) VALUES (?,?,?,?,?)',(kind,name,version,stamp,json.dumps(p,ensure_ascii=False,allow_nan=False)))
        p['id']=cur.lastrowid
    if kind in ['rule','template']:
        folder=ROOT/('rules' if kind=='rule' else 'templates')/'versions'
        folder.mkdir(exist_ok=True)
        (folder/f'{p["id"]}-v{version}.json').write_text(json.dumps(p,ensure_ascii=False,indent=2),encoding='utf-8')
    return p
for b in BUSINESSES:
    if not records('rule',b):
        insert('rule',b,{'business':b,'effective':'2026-10-01','approved':False,'notice':'Quy tắc khởi tạo, cần người phụ trách nghiệp vụ hiệu chỉnh/duyệt.','requireSource':True,'highPriceIsError':False,'requireEquivalentScope':True,'checks':['Đơn vị, hãng/model, thông số','VAT, vận chuyển, phụ kiện','Phạm vi cung cấp/thi công và thời điểm','Không tự chọn giá thấp nhất','Chỉ tính thành tiền theo cách người dùng xác nhận']})
    if not records('template',b):
        insert('template',b,{'business':b,'title':f'Báo cáo soát xét giá – {b}','columns':list(range(15)),'sections':['Cơ sở và phạm vi soát xét',f'Kết quả {b}','Các nội dung cần xác nhận']})

app=FastAPI(title='CCM-AI – Soát xét giá',docs_url=None,redoc_url=None)
class LocalSecurity(BaseHTTPMiddleware):
    async def dispatch(self,request,call_next):
        if request.headers.get('host','').split(':')[0].lower() not in ['localhost','127.0.0.1']:
            return FileResponse(ROOT/'frontend'/'dist'/'denied.txt',status_code=403) if (ROOT/'frontend'/'dist'/'denied.txt').exists() else __import__('starlette.responses',fromlist=['JSONResponse']).JSONResponse({'detail':'Local host only'},status_code=403)
        origin=request.headers.get('origin')
        if origin and origin!='https://localhost:3443':
            return __import__('starlette.responses',fromlist=['JSONResponse']).JSONResponse({'detail':'Origin denied'},status_code=403)
        if request.url.path.startswith('/api/') and request.url.path not in ['/api/session','/api/health']:
            if not secrets.compare_digest(request.headers.get('X-CCM-Token',''),TOKEN):
                return __import__('starlette.responses',fromlist=['JSONResponse']).JSONResponse({'detail':'Session required'},status_code=401)
        length=int(request.headers.get('content-length','0'))
        if length>MAX_BYTES+65536:
            return __import__('starlette.responses',fromlist=['JSONResponse']).JSONResponse({'detail':'Tệp tối đa 25 MB'},status_code=413)
        response=await call_next(request)
        response.headers['X-Content-Type-Options']='nosniff'
        response.headers['Cache-Control']='no-store' if request.url.path.startswith('/api') else 'no-cache'
        return response
app.add_middleware(LocalSecurity)
@app.get('/api/health')
def health(request:Request): return {'app':'CCM-AI','version':'0.3.0','installPath':str(ROOT),'ai':'ollama-local' if (ROOT/'config/ai.json').exists() else 'disabled','https':request.url.scheme=='https','storage':'SQLite','externalDataTransmission':False}
@app.get('/api/session')
def session(): return {'token':TOKEN}
@app.post('/api/acceptance')
def acceptance_receipt(body:dict):
    atomic_json(ROOT/'docs'/'local-officejs-acceptance.json',body)
    return {'saved':True}
def business(value):
    if value not in BUSINESSES: raise HTTPException(400,'Nghiệp vụ không hợp lệ')
    return value
@app.get('/api/rules')
def get_rule(business:str): return records('rule',globals()['business'](business))[-1]
@app.post('/api/rules')
def save_rule(body:dict):
    b=business(body.get('business'))
    if body.get('highPriceIsError') is not False or body.get('requireSource') is not True:
        raise HTTPException(400,'Không được coi giá cao tự động là lỗi; mỗi giá phải có nguồn.')
    if not isinstance(body.get('checks'),list) or not body.get('effective'): raise HTTPException(400,'Thiếu checks hoặc ngày áp dụng')
    return insert('rule',b,body)
@app.get('/api/templates')
def get_template(business:str): return records('template',globals()['business'](business))[-1]
@app.post('/api/templates')
def save_template(body:dict):
    b=business(body.get('business'));cols=body.get('columns',[])
    if not cols or len(set(cols))!=len(cols) or any(type(x)!=int or x<0 or x>14 for x in cols): raise HTTPException(400,'columns cần các vị trí 0–14, không trùng')
    if not isinstance(body.get('sections'),list): raise HTTPException(400,'sections phải là danh sách mục')
    return insert('template',b,body)
@app.get('/api/profiles')
def profiles(): return records('profile')
@app.post('/api/profiles')
def save_profile(body:dict):
    if not body.get('name') or len(body.get('maps',[]))!=2: raise HTTPException(400,'Thiếu mẫu ánh xạ hai vùng')
    return insert('profile',body['name'],body)
@app.get('/api/prices')
def prices(): return records('price')
@app.post('/api/prices')
def add_prices(body:dict):
    entries=body.get('entries',[])
    if not isinstance(entries,list): raise HTTPException(400,'entries phải là danh sách')
    for p in entries:
        if not all(p.get(k) for k in ['name','unit','source']) or type(p.get('price')) not in [float,int] or not math.isfinite(p['price']) or p['price']<0: raise HTTPException(400,'Mỗi giá cần name, unit, price là số >=0 và source')
    return [insert('price',p.get('code',p['name'])+'|'+p['source'],p) for p in entries]
@app.get('/api/documents')
def documents(): return [{k:v for k,v in d.items() if k!='text'} for d in records('document')]
@app.get('/api/documents/{doc_id}')
def document(doc_id:int):
    d=next((d for d in records('document') if d['id']==doc_id),None)
    if not d: raise HTTPException(404,'Không có tài liệu')
    return d
def extract(path:Path):
    if path.suffix=='.pdf':
        from pypdf import PdfReader
        r=PdfReader(path);return '\n'.join(f'[Trang {i+1}]\n'+(p.extract_text() or '') for i,p in enumerate(r.pages[:100]))[:300000]
    if path.suffix=='.docx':
        from docx import Document
        d=Document(path);return '\n'.join([p.text for p in d.paragraphs]+[' | '.join(c.text for c in row.cells) for t in d.tables for row in t.rows])[:300000]
    from openpyxl import load_workbook
    wb=load_workbook(path,read_only=True,data_only=False)
    out=[]
    for s in wb:
        out.append('[Sheet '+s.title+']')
        for row in s.iter_rows(max_row=min(s.max_row,5000),max_col=min(s.max_column,100),values_only=True):
            out.append(' | '.join(str(v) if v is not None else '' for v in row))
            if sum(map(len,out))>300000: break
    wb.close();return '\n'.join(out)[:300000]
@app.post('/api/documents')
async def upload(file:UploadFile=File(...),category:str=Form(...)):
    if category not in ['review','reference','guidance']: raise HTTPException(400,'Loại tài liệu không hợp lệ')
    suffix=Path(file.filename or '').suffix.lower()
    if suffix not in ['.xlsx','.docx','.pdf']: raise HTTPException(400,'Chỉ hỗ trợ xlsx/docx/pdf, không nhận macro')
    raw=await file.read(MAX_BYTES+1)
    if len(raw)>MAX_BYTES: raise HTTPException(413,'Tệp tối đa 25 MB')
    name=Path(file.filename.replace('\\','/')).name
    path=ROOT/'uploads'/(str(uuid.uuid4())+suffix);path.write_bytes(raw)
    try: text=extract(path)
    except Exception as e: text=f'Chưa trích được nội dung: {type(e).__name__}; cần kiểm tra file/OCR.'
    return insert('document',category+'|'+name,{'name':name,'category':category,'path':str(path.relative_to(ROOT)),'hash':hashlib.sha256(raw).hexdigest(),'text':text})
def backup_path(bid):
    try: uuid.UUID(bid)
    except ValueError: raise HTTPException(400,'ID bản sao không hợp lệ')
    return ROOT/'backups'/(bid+'.json')
def atomic_json(path,payload):
    temp=path.with_suffix('.tmp');temp.write_text(json.dumps(payload,ensure_ascii=False,allow_nan=False),encoding='utf-8');os.replace(temp,path)
@app.post('/api/backups')
def backup(body:dict):
    p=backup_path(body['id'])
    if p.exists(): raise HTTPException(409,'Bản sao đã tồn tại, không ghi đè')
    if not body.get('workbook') or not body.get('before'): raise HTTPException(400,'Bản sao thiếu phạm vi')
    atomic_json(p,body)
    with connect() as con: con.execute('INSERT INTO backup_index VALUES (?,?,?,?)',(body['id'],body['workbook'],'prepared',now()))
    return {'saved':True,'id':body['id']}
@app.get('/api/backups')
def backup_list(workbook:str):
    with connect() as con: return [dict(x) for x in con.execute('SELECT * FROM backup_index WHERE workbook=? ORDER BY updated DESC',(workbook,)).fetchall()]
@app.get('/api/backups/{bid}')
def load_backup(bid:str):
    p=backup_path(bid)
    if not p.exists(): raise HTTPException(404,'Không tìm thấy bản sao')
    return json.loads(p.read_text(encoding='utf-8'))
@app.post('/api/backups/{bid}/state')
def backup_state(bid:str,body:dict):
    if body.get('state') not in ['prepared','applied','restored','rolled_back']: raise HTTPException(400,'Trạng thái không hợp lệ')
    p=load_backup(bid);p['state']=body['state'];atomic_json(backup_path(bid),p)
    with connect() as con: con.execute('UPDATE backup_index SET state=?,updated=? WHERE id=?',(body['state'],now(),bid))
    return {'ok':True}
@app.post('/api/reports')
def report(body:dict):
    from backend.export import export_report
    fmt=body.get('format');cols=body.get('columns')
    if fmt not in ['xlsx','docx','pdf']: raise HTTPException(400,'Định dạng không hỗ trợ')
    if not isinstance(cols,list) or not cols or any(type(c)!=int or not 0<=c<15 for c in cols) or len(set(cols))!=len(cols): raise HTTPException(400,'Cột không hợp lệ')
    if len(body.get('headers',[]))!=15 or any(len(r)!=15 for r in body.get('rows',[])): raise HTTPException(400,'Báo cáo cần 15 cột dữ liệu chuẩn')
    business(body.get('business'))
    filename=f'CCM-{datetime.now().strftime("%Y%m%d-%H%M%S")}-{uuid.uuid4().hex[:6]}.{fmt}'
    path=ROOT/'reports'/filename
    export_report(path,body)
    atomic_json(path.with_suffix('.json'),body)
    return {'filename':filename,'url':'/api/reports/'+filename}
@app.get('/api/reports/{name}')
def download_report(name:str):
    if Path(name).name!=name or not name.startswith('CCM-'): raise HTTPException(400,'Tên báo cáo không hợp lệ')
    p=ROOT/'reports'/name
    if not p.is_file(): raise HTTPException(404,'Không có báo cáo')
    return FileResponse(p,filename=name)
from backend.local_ai import make_router
app.include_router(make_router(ROOT,records,insert,business))
DIST=ROOT/'frontend'/'dist'
if DIST.exists(): app.mount('/',StaticFiles(directory=DIST,html=True),name='frontend')
