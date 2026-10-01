from pathlib import Path
from zipfile import ZipFile
ROOT=Path(__file__).resolve().parents[1]
target=Path(r'D:\00.Setupgpt\SoatXetGia')
if target.exists():raise RuntimeError('Đường dẫn đích đã tồn tại; không ghi đè.')
with ZipFile(ROOT/'CCM-AI-Ban-giao-0.1.zip') as archive:
 entries=[]
 for info in archive.infolist():
  if not info.filename.startswith('SoatXetGia/'):raise RuntimeError('Tiền tố archive không hợp lệ')
  relative=Path(info.filename[len('SoatXetGia/'):]);dest=(target/relative).resolve()
  if not dest.is_relative_to(target.resolve()):raise RuntimeError('Đường dẫn vượt thư mục đích')
  entries.append((info,dest))
 target.mkdir(parents=True,exist_ok=False)
 for info,dest in entries:
  if info.is_dir():dest.mkdir(parents=True,exist_ok=True)
  else:
   dest.parent.mkdir(parents=True,exist_ok=True)
   with dest.open('xb') as out:out.write(archive.read(info))
print(str(target))
