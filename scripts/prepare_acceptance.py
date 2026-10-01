"""Adapt Microsoft's sideload test fixture; do not modify a user's workbook."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import urllib.request
ROOT=Path(__file__).resolve().parents[1]
url='https://raw.githubusercontent.com/OfficeDev/Office-Addin-Scripts/master/packages/office-addin-dev-settings/templates/ExcelWorkbookWithTaskPane.xlsx'
source=ROOT/'tests'/'microsoft-sideload-template.xlsx'
if not source.exists():urllib.request.urlretrieve(url,source)
dest=ROOT/'tests'/'CCM-AI-Acceptance.xlsx'
if dest.exists():raise RuntimeError('Test workbook already exists; do not overwrite it.')
with ZipFile(source) as z,ZipFile(dest,'w',ZIP_DEFLATED) as out:
 for name in z.namelist():
  data=z.read(name)
  if name=='xl/webextensions/webextension.xml':data=data.replace(b'00000000-0000-0000-0000-000000000000',b'79e68a1b-4a3b-41d2-970b-f2d83c16dc01').replace(b'1.0.0.0',b'0.1.0.0')
  if name=='docProps/core.xml':
   from xml.etree import ElementTree as ET
   el=ET.fromstring(data);ns='http://purl.org/dc/elements/1.1/';title=el.find('{'+ns+'}title')
   if title is None:title=ET.SubElement(el,'{'+ns+'}title')
   title.text='CCM-AI-ACCEPTANCE-ONLY';data=ET.tostring(el,encoding='utf-8',xml_declaration=True)
  out.writestr(name,data)
print(dest)
