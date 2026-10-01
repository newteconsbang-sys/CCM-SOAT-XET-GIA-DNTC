"""Extract CPU members from the pinned official release's trailing ZIP range."""
from pathlib import Path
from zipfile import ZipFile
import sys
archive,target=Path(sys.argv[1]),Path(sys.argv[2]).resolve()
target.mkdir(parents=True,exist_ok=True)
with ZipFile(archive) as z:
    for entry in z.infolist():
        if entry.header_offset<0 or entry.is_dir() or '/vulkan/' in entry.filename: continue
        destination=(target/entry.filename).resolve()
        if not destination.is_relative_to(target): raise ValueError('Unsafe archive path')
        if not (entry.filename=='ollama.exe' or entry.filename.startswith('lib/ollama/')): raise ValueError('Unexpected archive member')
        destination.parent.mkdir(parents=True,exist_ok=True)
        destination.write_bytes(z.read(entry)) # ZipFile checks each member's CRC.
print('Installed Ollama CPU runtime')
