"""Local file intake: preserve originals, map columns, preview, then version prices."""
import csv
import hashlib
import io
import json
import math
import re
import uuid
import zipfile
from pathlib import Path
from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from openpyxl import load_workbook
from .pdf_intake import extract_pdf, pdf_price, excel_bytes
from .price_expressions import expression, expression_conditions
from fastapi.responses import FileResponse
from fastapi.responses import Response

ROLES = ('description', 'unit', 'brand', 'material', 'labor', 'source', 'date', 'conditions', 'quantity')
LIMIT_BYTES = 25 * 1024 * 1024
LIMIT_ROWS = 10000
LIMIT_COLS = 100
LIMIT_CELLS = 500000

def fail(message):
    raise HTTPException(400, message)

def parse_file(raw, suffix):
    if suffix == '.pdf':
        return extract_pdf(raw)[0], None
    if suffix == '.xlsx':
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                if sum(info.file_size for info in archive.infolist()) > 100 * 1024 * 1024:
                    fail('Excel giải nén vượt 100 MB. Chọn file nhỏ hơn.')
                if any('vbaproject' in name.lower() for name in archive.namelist()):
                    fail('Chỉ nhận XLSX không có macro.')
            workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=False, keep_links=False)
            sheets = {}
            cell_count = 0
            try:
                for sheet in workbook.worksheets:
                    if sheet.sheet_state != 'visible':
                        continue
                    if len(sheets) >= 30 or sheet.max_row > LIMIT_ROWS or sheet.max_column > LIMIT_COLS:
                        fail('File tối đa 30 sheet hiện, 10.000 dòng và 100 cột mỗi sheet. Tách vùng giá trước khi nạp.')
                    cell_count += sheet.max_row * sheet.max_column
                    if cell_count > LIMIT_CELLS:
                        fail('File vượt 500.000 ô. Tách vùng thư viện cần nạp thành file riêng.')
                    def source_value(cell):
                        value = cell.value
                        # Excel stores 5% as 0.05: preserve its meaning rather than a 0.05 VND price.
                        fmt = re.sub(r'"[^"]*"|\\.', '', cell.number_format or '')
                        if type(value) in (int, float) and '%' in fmt:
                            from decimal import Decimal
                            display = format(Decimal(str(value)) * 100, 'f')
                            return (display.rstrip('0').rstrip('.') if '.' in display else display) + '%'
                        return value.isoformat() if hasattr(value, 'isoformat') else value
                    rows = [[source_value(cell) for cell in row] for row in sheet.iter_rows()]
                    sheets[sheet.title] = rows
            finally:
                workbook.close()
            if not sheets:
                fail('File không có sheet hiện để nạp.')
            return sheets, None
        except HTTPException:
            raise
        except Exception:
            fail('Không đọc được XLSX. File cần không có mật khẩu và được lưu đúng định dạng.')
    if suffix == '.csv':
        try:
            content = raw.decode('utf-8-sig')
        except UnicodeError:
            fail('CSV cần lưu UTF-8. Có thể dùng XLSX để giữ tiếng Việt.')
        try:
            dialect = csv.Sniffer().sniff(content[:8192], delimiters=',;\t')
        except csv.Error:
            dialect = csv.excel
        rows = []
        cell_count = 0
        for row in csv.reader(io.StringIO(content), dialect):
            if len(rows) >= LIMIT_ROWS or len(row) > LIMIT_COLS:
                fail('CSV tối đa 10.000 dòng và 100 cột.')
            cell_count += len(row)
            if cell_count > LIMIT_CELLS:
                fail('CSV vượt 500.000 ô. Tách file trước khi nạp.')
            rows.append(row)
        return {'CSV': rows}, None
    if suffix == '.json':
        try:
            document = json.loads(raw.decode('utf-8-sig'))
        except (ValueError, UnicodeError):
            fail('JSON không hợp lệ.')
        selected = document.get('business') if isinstance(document, dict) else None
        entries = document.get('entries') if isinstance(document, dict) else document
        if not isinstance(entries, list) or not 1 <= len(entries) <= 5000 or any(not isinstance(p, dict) for p in entries):
            fail('JSON cần danh sách 1–5.000 đầu mục hoặc object có entries.')
        return {'JSON': [list(ROLES)] + [[p.get(key) if p.get(key) is not None else p.get(key+'Expression') if key in ('material','labor') else None for key in ROLES] for p in entries]}, selected
    fail('Chọn XLSX, PDF, CSV hoặc JSON. File XLS cũ cần lưu lại thành XLSX.')

def numeric_price(value, row, role):
    if value is None or isinstance(value, str) and not value.strip():
        return None
    if isinstance(value, str):
        value = value.strip()
        if not re.fullmatch(r'\d+(?:\.\d+)?', value):
            fail(f'Dòng {row}, {role}: cần giá số không âm; CSV/JSON dùng số không có phân cách hàng nghìn. Công thức cần chuyển thành giá trị trước.')
        value = float(value)
    if type(value) not in (int, float) or not math.isfinite(value) or value < 0:
        fail(f'Dòng {row}, {role}: giá không hợp lệ.')
    return value

def mapped_entries(rows, columns, header_row, last_row, filename, sheet, digest, pdf_locations=None, price_format=None, quantity_format=None, unsupported_prices=None):
    if not isinstance(columns, dict) or set(columns) - set(ROLES):
        fail('Ánh xạ cột không hợp lệ.')
    indices = [value for value in columns.values() if value is not None]
    if any(type(value) is not int or not 0 <= value < LIMIT_COLS for value in indices) or len(indices) != len(set(indices)):
        fail('Các vai trò phải dùng cột khác nhau.')
    if columns.get('description') is None or columns.get('unit') is None or columns.get('material') is None and columns.get('labor') is None:
        fail('Chọn mô tả, ĐVT và ít nhất một cột giá.')
    if columns.get('quantity') is not None and quantity_format is not None and quantity_format not in ('vi','en','plain'):
        fail('Chọn định dạng số cho khối lượng.')
    if type(header_row) is not int or type(last_row) is not int or not 1 <= header_row < last_row <= len(rows):
        fail('Chọn dòng tiêu đề và dòng cuối hợp lệ. Dữ liệu bắt đầu sau dòng tiêu đề.')
    entries, skipped = [], []
    for index in range(header_row, last_row):
        row = rows[index]
        def cell(role):
            column = columns.get(role)
            return row[column] if column is not None and column < len(row) else None
        def text(role):
            value = cell(role)
            if value is None:
                return ''
            if isinstance(value, (dict, list)) or isinstance(value, str) and len(value) > 2000:
                fail(f'Dòng {index+1}: nội dung cột {role} quá dài hoặc sai kiểu.')
            return str(value).strip()
        description, unit = text('description'), text('unit')
        source_expressions = {role+'Expression': expression(cell(role)) for role in ('material','labor') if expression(cell(role))}
        if not description or not unit and not source_expressions or unit.casefold() in ['-', 'đvt', 'đơn vị', 'đơn vị tính', 'unit']:
            skipped.append(index + 1)
            continue
        row_issues = []
        def price(role, label):
            value = cell(role)
            if role in ('material','labor') and expression(value):
                if len(value.strip()) > 2000:fail(f'Dòng {index+1}: tỷ lệ / biểu thức vượt 2.000 ký tự.')
                issue=dict(row=index+1,role=label,value=value.strip())
                row_issues.append(issue)
                if unsupported_prices is not None:unsupported_prices.append(issue)
                return None
            selected_format = (quantity_format or price_format) if role=='quantity' else price_format
            if pdf_locations is not None or selected_format in ('vi','en','plain'):
                value = pdf_price(value, index + 1, label, selected_format)
            return numeric_price(value, index + 1, label)
        material = price('material', 'vật tư')
        labor = price('labor', 'nhân công')
        quantity = price('quantity', 'khối lượng')
        if material is None and labor is None and not source_expressions:
            skipped.append(index + 1)
            continue
        source = f'{filename} | {sheet}!dòng {index+1} | SHA256 {digest}'
        if pdf_locations is not None:
            ref = pdf_locations[index]
            source += f' | PDF trang {ref["page"]}, bảng {ref["table"]}, dòng {ref["row"]}'
        supplied_source = text('source')
        if supplied_source:
            source += ' | ' + supplied_source
        conditions = expression_conditions(dict(conditions=text('conditions'), unit=unit, **source_expressions))
        entries.append(dict(description=description, unit=unit, quantity=quantity, brand=text('brand'), material=material, labor=labor,
                            source=source[:2000], date=text('date'), conditions=conditions, **source_expressions))
        if len(entries) > 5000:
            fail('Mỗi phiên bản thư viện tối đa 5.000 đầu mục.')
    if not entries:
        fail('Không có dòng có mô tả và giá số hoặc tỷ lệ / biểu thức. Kiểm tra dòng tiêu đề và các cột.')
    return entries, skipped

def make_router(storage_root, business, save_library):
    root = Path(storage_root) / 'uploads' / 'price-library'
    router = APIRouter(prefix='/intake')

    def draft(identifier):
        if not isinstance(identifier, str) or not re.fullmatch('[a-f0-9]{32}', identifier):
            fail('Mã file nạp không hợp lệ.')
        folder = root / identifier
        if not (folder / 'metadata.json').exists():
            raise HTTPException(404, 'Không tìm thấy file đã nạp.')
        meta = json.loads((folder / 'metadata.json').read_text(encoding='utf-8'))
        raw = (folder / ('original' + meta['suffix'])).read_bytes()
        if hashlib.sha256(raw).hexdigest() != meta['sha256']:
            raise HTTPException(409, 'File gốc đã thay đổi; nạp và xem trước lại.')
        if meta['suffix'] == '.pdf':
            extracted = (folder / 'extracted.json').read_bytes()
            if hashlib.sha256(extracted).hexdigest() != meta['extractedHash']:
                raise HTTPException(409, 'Bảng chuyển PDF đã thay đổi; nạp và xem trước lại.')
            sheets = json.loads(extracted)
        else:
            sheets, selected = parse_file(raw, meta['suffix'])
        return meta, sheets

    @router.post('/{identifier}/recognize')
    def recognize_file(identifier:str, body:dict):
        from .recognize_columns import recognize
        meta, sheets = draft(identifier)
        return recognize(sheets, storage_root, body.get('sheet'), body.get('headerRow'))

    @router.post('/upload')
    async def upload(file:UploadFile=File(...), selected_business:str=Form(...)):
        selected_business = business(selected_business)
        name = Path((file.filename or '').replace('\\', '/')).name[:150]
        suffix = Path(name).suffix.lower()
        raw = await file.read(LIMIT_BYTES + 1)
        if not raw or len(raw) > LIMIT_BYTES:
            fail('File phải có dữ liệu và tối đa 25 MB.')
        extraction = None
        if suffix == '.pdf':
            sheets, extraction = extract_pdf(raw)
            source_business = None
        else:
            sheets, source_business = parse_file(raw, suffix)
        code = 'DNTC' if selected_business == 'Điện nước ĐNTC' else selected_business
        if source_business is not None and source_business not in [code, selected_business, 'ĐNTC' if code == 'DNTC' else code]:
            fail('Nghiệp vụ trong JSON khác thư viện đang chọn; chọn lại nghiệp vụ.')
        identifier = uuid.uuid4().hex
        meta = dict(id=identifier, filename=name, suffix=suffix, business=selected_business, sha256=hashlib.sha256(raw).hexdigest())
        folder = root / identifier
        folder.mkdir(parents=True)
        (folder / ('original' + suffix)).write_bytes(raw)
        if extraction is not None:
            extracted = json.dumps(sheets, ensure_ascii=False, allow_nan=False).encode('utf-8')
            meta['extractedHash'] = hashlib.sha256(extracted).hexdigest()
            meta['extraction'] = extraction
            (folder / 'extracted.json').write_bytes(extracted)
            (folder / 'converted.xlsx').write_bytes(excel_bytes(sheets, extraction['locations'], name, meta['sha256']))
        (folder / 'metadata.json').write_text(json.dumps(meta, ensure_ascii=False), encoding='utf-8')
        return dict({k:v for k,v in meta.items() if k != 'extraction'}, extraction={k:v for k,v in extraction.items() if k != 'locations'} if extraction else None,
                    sheets=[dict(name=key, rows=len(rows), columns=max((len(row) for row in rows), default=0)) for key, rows in sheets.items()])

    @router.get('/{identifier}/converted.xlsx')
    def converted(identifier:str):
        meta, sheets = draft(identifier)
        if meta['suffix'] != '.pdf':
            fail('Chỉ PDF có bản Excel chuyển đổi.')
        # Rebuild from checked extraction rather than trusting a modified download file.
        path = root / identifier / 'converted.xlsx'
        path.write_bytes(excel_bytes(sheets, meta['extraction']['locations'], meta['filename'], meta['sha256']))
        return FileResponse(path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=Path(meta['filename']).stem+'-chuyen-PDF.xlsx')

    @router.get('/{identifier}/sheet')
    def read_sheet(identifier:str, sheet:str, header_row:int=1):
        meta, sheets = draft(identifier)
        if sheet not in sheets or not 1 <= header_row < len(sheets[sheet]):
            fail('Sheet hoặc dòng tiêu đề không hợp lệ.')
        rows = sheets[sheet]
        return dict(sheet=sheet, rows=len(rows), headers=rows[header_row-1], preview=rows[header_row:header_row+12])

    def mapped(body):
        meta, sheets = draft(body.get('id'))
        selected_business = business(body.get('business'))
        if selected_business != meta['business']:
            fail('File đã nạp thuộc nghiệp vụ khác. Nạp lại vào đúng thư viện.')
        sheet = body.get('sheet')
        if sheet not in sheets:
            fail('Chọn sheet hợp lệ.')
        refs = meta.get('extraction', {}).get('locations', {}).get(sheet) if meta['suffix'] == '.pdf' else None
        unsupported = []
        entries, skipped = mapped_entries(sheets[sheet], body.get('columns'), body.get('headerRow'), body.get('lastRow'), meta['filename'], sheet, meta['sha256'], refs, body.get('priceFormat'), body.get('quantityFormat'), unsupported)
        return meta, entries, skipped, unsupported

    @router.post('/preview')
    def preview(body:dict):
        meta, entries, skipped, unsupported = mapped(body)
        return dict(count=len(entries), entries=entries[:100], skippedRows=skipped, unsupportedPrices=unsupported[:100], unsupportedPriceCount=len(unsupported), filename=meta['filename'], sha256=meta['sha256'])

    @router.post('/export')
    def export_five_columns(body:dict):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        meta, entries, skipped, unsupported = mapped(body)
        book=Workbook();sheet=book.active;sheet.title='Thư viện giá'
        sheet.append(['Hạng mục','Đơn vị','Khối lượng','Đơn giá vật tư','Đơn giá nhân công'])
        for entry in entries:
            sheet.append([entry['description'],entry['unit'],entry.get('quantity'),entry.get('materialExpression') or entry['material'],entry.get('laborExpression') or entry['labor']])
        for row in sheet:
            for cell in row:
                if isinstance(cell.value,str):cell.data_type='s'
                cell.alignment=Alignment(vertical='top',wrap_text=True)
                if cell.row==1:
                    cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='087F79')
                elif cell.column>=3:cell.number_format='#,##0.####'
        for column,width in zip('ABCDE',[45,12,16,23,23]):sheet.column_dimensions[column].width=width
        sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
        book.properties.description=f'CCM-AI; file gốc: {meta["filename"]}; SHA256 {meta["sha256"]}; chỉ xuất dữ liệu đã ánh xạ, không tự tạo giá.'
        output=io.BytesIO();book.save(output)
        return Response(output.getvalue(),media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename="CCM-Thu-vien-5-cot.xlsx"'})

    @router.post('/commit')
    def commit(body:dict):
        if body.get('confirmed') is not True:
            fail('Xem trước và xác nhận trước khi lưu thư viện.')
        meta, entries, skipped, unsupported = mapped(body)
        if body.get('sourceHash') != meta['sha256']:
            raise HTTPException(409, 'Nguồn đã đổi; nạp và xem trước lại.')
        if meta['suffix'] == '.pdf' and body.get('pdfReviewed') is not True:
            fail('Đối chiếu PDF gốc và xác nhận đã kiểm tra trước khi lưu.')
        saved = save_library(dict(business=meta['business'], name=body.get('name') or meta['filename'], entries=entries, confirmed=True))
        return dict(saved, skippedRows=skipped, unsupportedPrices=unsupported[:100], unsupportedPriceCount=len(unsupported))

    return router
