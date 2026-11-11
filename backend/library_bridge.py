"""Read-only, versioned price snapshot for the Excel/ChatGPT bridge."""
import hashlib
import json
import math
from datetime import datetime, timezone
from fastapi import HTTPException
from .price_expressions import expression_fields, expression_conditions

BUSINESS_CODES = {'Điện nước ĐNTC': 'DNTC', 'EW': 'EW', 'NCC': 'NCC', 'NTP': 'NTP'}
FIELDS = ['description', 'brand', 'unit', 'material', 'labor', 'source', 'conditions', 'date']

def snapshot(document, business):
    code = BUSINESS_CODES[business]
    entries = document.get('entries')
    if not entries:
        raise HTTPException(409, f'Thư viện {code} chưa có giá. Nạp và lưu thư viện trong CCM-AI trước.')
    if not isinstance(entries, list) or len(entries) > 5000 or document.get('business', business) not in [business, code]:
        raise HTTPException(409, 'Thư viện không hợp lệ hoặc sai nghiệp vụ; cần nạp lại.')
    clean = []
    for entry in entries:
        if not isinstance(entry, dict):
            raise HTTPException(409, 'Dòng thư viện không hợp lệ.')
        row = {key: entry.get(key, None if key in ['material', 'labor'] else '') for key in FIELDS}
        try:expressions = expression_fields(entry)
        except ValueError as error:raise HTTPException(409,str(error))
        if expressions:
            row.update(expressions)
            row['conditions'] = expression_conditions(row)
        if any(not isinstance(row[key], str) or len(row[key]) > 32767 for key in FIELDS if key not in ['material', 'labor']):
            raise HTTPException(409, 'Nội dung thư viện không hợp lệ cho Excel.')
        if any(not row[key].strip() for key in ['description', 'source']) or not row['unit'].strip() and not expressions:
            raise HTTPException(409, 'Thư viện thiếu mô tả, ĐVT hoặc nguồn.')
        prices = [row['material'], row['labor']]
        if all(value is None for value in prices) and not expressions or any(value is not None and (type(value) not in [int, float] or not math.isfinite(value) or value < 0) for value in prices):
            raise HTTPException(409, 'Thư viện có đơn giá không hợp lệ.')
        clean.append(row)
        if 'quantity' in entry:
            quantity=entry['quantity']
            if quantity is not None and (type(quantity) not in [int,float] or not math.isfinite(quantity) or quantity<0):
                raise HTTPException(409,'Khối lượng thư viện không hợp lệ.')
            row['quantity']=quantity
    version = str(document.get('version', ''))
    if not version or version == '0':
        raise HTTPException(409, 'Thư viện thiếu phiên bản.')
    content = {'business': code, 'version': version, 'entries': clean}
    digest = hashlib.sha256(json.dumps(content, ensure_ascii=False, sort_keys=True, separators=(',', ':'), allow_nan=False).encode('utf-8')).hexdigest()
    return dict(content, schema='ccm-ai-library-v1', sha256=digest,
                exportedAt=datetime.now(timezone.utc).isoformat(), count=len(clean))
