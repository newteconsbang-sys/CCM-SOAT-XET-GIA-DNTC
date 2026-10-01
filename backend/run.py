from pathlib import Path
import sys,json
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'backend'/'vendor'))
sys.path.insert(0,str(ROOT))
if __name__=='__main__':
    import uvicorn
    uvicorn.run('backend.app:app',host='127.0.0.1',port=3443,ssl_keyfile=str(ROOT/'config'/'localhost.key'),ssl_certfile=str(ROOT/'config'/'localhost.crt'),log_level='warning')
