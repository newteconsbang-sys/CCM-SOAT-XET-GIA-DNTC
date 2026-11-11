"""Read prices uploaded to the local CCM-AI web. Never writes libraries or Excel."""
import argparse
import json
import sys
from pathlib import Path

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--app-root',type=Path,default=Path(r'D:\00.Setupgpt\SoatXetGia'))
    parser.add_argument('--action',choices=['list','read','review'],required=True)
    parser.add_argument('--business',choices=['DNTC','EW','NCC','NTP'])
    parser.add_argument('--version')
    parser.add_argument('--offset',type=int,default=0)
    parser.add_argument('--limit',type=int,default=200)
    parser.add_argument('--rows',type=Path,help='JSON rows with Excel row, description, unit, brand')
    args=parser.parse_args()
    sys.path.insert(0,str(args.app_root))
    try:
        from backend.price_service import LibraryService
        service=LibraryService(args.app_root/'libraries')
        if args.action=='list':
            result=service.call('ccm_list_libraries',{})
        else:
            if not args.business:parser.error('--business required')
            arguments={'business':args.business}
            if args.version:arguments['version']=args.version
            if args.action=='read':
                arguments.update(offset=args.offset,limit=args.limit)
                result=service.call('ccm_read_library',arguments)
            else:
                if not args.rows or not args.version:parser.error('review requires --rows and --version')
                if args.rows.stat().st_size>2_000_000:raise ValueError('Rows JSON exceeds 2 MB')
                arguments['rows']=json.loads(args.rows.read_text(encoding='utf-8-sig'))
                result=service.call('ccm_review_rows',arguments)
        print(json.dumps(result,ensure_ascii=False,allow_nan=False))
    except (ImportError,ValueError,KeyError,TypeError,OSError) as error:
        print(json.dumps({'error':str(error),'library_connected':False},ensure_ascii=False))
        return 1
    return 0

if __name__=='__main__':
    if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
