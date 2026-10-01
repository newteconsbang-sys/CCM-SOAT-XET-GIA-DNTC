"""Import a versioned CCM library. Does not edit a workbook."""
import argparse
import csv
import json
import math
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

BUSINESSES = ('DNTC', 'EW', 'NCC', 'NTP')
FIELDS = ('description', 'unit', 'brand', 'material', 'labor', 'source', 'date', 'conditions')

def initialize(root):
    for business in BUSINESSES:
        folder = root / business
        (folder / 'versions').mkdir(parents=True, exist_ok=True)
        template = folder / 'Mau-thu-vien.csv'
        if not template.exists():
            with template.open('w', encoding='utf-8-sig', newline='') as stream:
                csv.writer(stream).writerow(FIELDS)

def price(value):
    if value is None or value == '':
        return None
    if isinstance(value, bool) or (isinstance(value, str) and not re.fullmatch(r'\d+(?:\.\d+)?', value.strip())):
        raise ValueError('Price must be numeric without thousands separators')
    number = float(value)
    if not math.isfinite(number) or number < 0:
        raise ValueError('Price must be finite and nonnegative')
    return int(number) if number.is_integer() else number

def import_library(root, business, source):
    if business not in BUSINESSES:
        raise ValueError('Unknown business')
    if source.suffix.lower() == '.csv':
        with source.open(encoding='utf-8-sig', newline='') as stream:
            entries = list(csv.DictReader(stream))
    else:
        entries = json.loads(source.read_text(encoding='utf-8-sig'))
    if not isinstance(entries, list) or not entries:
        raise ValueError('Library must contain entries')
    clean = []
    for index, entry in enumerate(entries, 1):
        record = {key: entry.get(key, '') for key in FIELDS}
        for key in ('description', 'unit', 'brand', 'source', 'date', 'conditions'):
            record[key] = str(record[key] or '').strip()
        if not record['description'] or not record['unit']:
            raise ValueError(f'Entry {index}: description and unit required')
        record['material'], record['labor'] = price(record['material']), price(record['labor'])
        if record['material'] is None and record['labor'] is None:
            raise ValueError(f'Entry {index}: at least one price required')
        record['source'] = record['source'] or f'{source.name}: record {index}'
        clean.append(record)
    initialize(root)
    version = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ') + '-' + uuid.uuid4().hex[:12]
    folder = root / business
    (folder / 'versions' / (version + '.json')).write_text(json.dumps({'business': business, 'version': version, 'entries': clean}, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary = folder / ('active-' + uuid.uuid4().hex + '.tmp')
    temporary.write_text(json.dumps({'version': version}), encoding='utf-8')
    temporary.replace(folder / 'active.json')
    return {'business': business, 'version': version, 'entries': len(clean)}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path(r'D:\00.Setupgpt\SoatXetGia\libraries'))
    parser.add_argument('--init', action='store_true')
    parser.add_argument('--business', choices=BUSINESSES)
    parser.add_argument('--input', type=Path)
    args = parser.parse_args()
    if args.init:
        initialize(args.root)
        result = {'root': str(args.root), 'businesses': BUSINESSES}
    else:
        if not args.business or not args.input:
            parser.error('--business and --input required for import')
        result = import_library(args.root, args.business, args.input)
    print(json.dumps(result, ensure_ascii=True))

if __name__ == '__main__':
    main()
