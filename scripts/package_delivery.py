from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'CCM-AI-Ban-giao-0.1.zip'
allowed=['frontend/src','frontend/public','frontend/dist','backend','manifests','rules','templates','docs','scripts','tests','runtime']
exclude=['node_modules','__pycache__','.venv','.git','data','uploads','reports','backups','config']
files=[]
for folder in allowed:
 for p in (ROOT/folder).rglob('*'):
  rel=p.relative_to(ROOT)
  if p.is_file() and not any(x in exclude for x in rel.parts) and not p.name.startswith('local-') and not p.name.startswith('refine_') and p.suffix not in ['.pyc','.log']:files.append(p)
for name in ['README.md','.gitignore','Cai-dat-CCM.cmd','Khởi động CCM-AI.cmd','frontend/package.json','frontend/tsconfig.json','frontend/index.html','frontend/pnpm-lock.yaml','docs/YEU-CAU.txt']:
 p=ROOT/name
 if p.exists():files.append(p)
with ZipFile(OUT,'w',ZIP_DEFLATED,compresslevel=6) as z:
 for p in sorted(set(files)):z.write(p,'SoatXetGia/'+p.relative_to(ROOT).as_posix())
 for folder in ['data','uploads','reports','backups','config']:z.writestr('SoatXetGia/'+folder+'/.keep','')
checksum=hashlib.sha256(OUT.read_bytes()).hexdigest()
(ROOT/'CCM-AI-Ban-giao-0.1.sha256').write_text(checksum+'  '+OUT.name+'\n',encoding='ascii')
print(json.dumps({'file':str(OUT),'bytes':OUT.stat().st_size,'sha256':checksum,'files':len(set(files))}))
