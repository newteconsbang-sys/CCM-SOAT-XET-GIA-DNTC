"""MCP stdio: UTF-8 JSON-RPC, one object per line, stdout is protocol only."""
import json
import os
import sys
from library_service import LibraryService, rpc

def main():
    sys.stdin.reconfigure(encoding='utf-8')
    sys.stdout.reconfigure(encoding='utf-8')
    service = LibraryService(os.environ.get('CCM_LIBRARY_ROOT', r'D:\00.Setupgpt\SoatXetGia\libraries'))
    while True:
        line = sys.stdin.readline(2_000_001)
        if not line:
            break
        if len(line) > 2_000_000:
            print('MCP input exceeded limit', file=sys.stderr)
            return
        try:
            response = rpc(service, json.loads(line))
        except (ValueError, TypeError):
            response = {'jsonrpc': '2.0', 'id': None, 'error': {'code': -32700, 'message': 'Parse error'}}
        if response is not None:
            print(json.dumps(response, ensure_ascii=False, allow_nan=False), flush=True)

if __name__ == '__main__':
    main()
