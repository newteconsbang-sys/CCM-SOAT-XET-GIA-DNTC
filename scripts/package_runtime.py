"""Copy the required runtime without depending on an open Codex app."""
from pathlib import Path
import sys,shutil,importlib.util
ROOT=Path(__file__).resolve().parents[1]
base=Path(sys.base_prefix);dest=ROOT/'runtime'/'python';dest.mkdir(parents=True,exist_ok=True)
for name in ['python.exe','pythonw.exe','python3.dll','python312.dll','vcruntime140.dll','vcruntime140_1.dll','LICENSE.txt']:
 p=base/name
 if p.exists():shutil.copy2(p,dest/name)
for name in ['DLLs','Lib']:
 shutil.copytree(base/name,dest/name,dirs_exist_ok=True,ignore=shutil.ignore_patterns('site-packages','__pycache__','test','tests','idlelib','ensurepip','tkinter'))
site=dest/'Lib'/'site-packages';site.mkdir(exist_ok=True)
for name in ['cryptography','cffi','pycparser','openpyxl','et_xmlfile','docx','lxml','reportlab','PIL','pypdf','charset_normalizer','typing_extensions']:
 spec=importlib.util.find_spec(name)
 if spec is None:continue
 source=Path(spec.origin)
 if source.name=='__init__.py':shutil.copytree(source.parent,site/source.parent.name,dirs_exist_ok=True,ignore=shutil.ignore_patterns('__pycache__','tests'))
 else:shutil.copy2(source,site/source.name)
 # Preserve license/version metadata for redistributed packages.
 for metadata in (base/'Lib'/'site-packages').glob(name.replace('_','-')+'-*.dist-info'):shutil.copytree(metadata,site/metadata.name,dirs_exist_ok=True)
 for metadata in (base/'Lib'/'site-packages').glob(name+'-*.dist-info'):shutil.copytree(metadata,site/metadata.name,dirs_exist_ok=True)
for file in (base/'Lib'/'site-packages').glob('_cffi_backend*'):shutil.copy2(file,site/file.name)
print('Portable Python:',dest)
