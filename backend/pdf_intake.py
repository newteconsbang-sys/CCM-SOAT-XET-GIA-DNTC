"""Extract local PDF tables as reviewable data; never execute PDF content."""
import io
import re
import unicodedata
from fastapi import HTTPException


def fail(message):
    raise HTTPException(400, message)


def clean(value):
    return re.sub(r'\s+', ' ', unicodedata.normalize('NFC', str(value or ''))).strip()


def extract_pdf(raw):
    import pdfplumber
    from pdfminer.pdfdocument import PDFPasswordIncorrect
    sheets, locations, warnings, groups = {}, {}, [], {}
    count, cells = 0, 0
    try:
        with pdfplumber.open(io.BytesIO(raw), unicode_norm='NFC') as document:
            pages = len(document.pages)
            if pages > 50:
                fail('PDF tối đa 50 trang. Tách phần bảng giá cần nạp thành PDF riêng.')
            for number, page in enumerate(document.pages, 1):
                if not page.chars:
                    warnings.append(f'Trang {number}: không có chữ đọc được; có thể là ảnh scan, cần OCR. Chưa lấy giá từ trang này.')
                    page.close()
                    continue
                if len(page.chars) > 100000:
                    fail(f'Trang {number} quá nhiều chữ. Tách bảng giá trước khi nạp.')
                tables = page.extract_tables()
                if not tables:
                    tables = page.extract_tables(dict(vertical_strategy='text', horizontal_strategy='text', min_words_vertical=3))
                    if tables:
                        warnings.append(f'Trang {number}: suy ra cột từ vị trí chữ vì không có đường kẻ; kiểm tra kỹ các cột và dòng.')
                found = 0
                for index, table in enumerate(tables, 1):
                    rows = [[clean(value) for value in row] for row in table]
                    rows = [row for row in rows if any(row)]
                    width = max((len(row) for row in rows), default=0)
                    if len(rows) < 2 or width < 2:
                        continue
                    count += 1
                    cells += len(rows) * width
                    if count > 30 or len(rows) > 10000 or width > 100 or cells > 500000:
                        fail('PDF quá nhiều bảng/ô. Tối đa 30 bảng, 10.000 dòng và 100 cột mỗi bảng, 500.000 ô; tách phần giá cần nạp.')
                    found += 1
                    name = f'Trang {number} - Bang {index}'
                    sheets[name] = rows
                    locations[name] = [dict(page=number, table=index, row=i+1) for i in range(len(rows))]
                    key = tuple(cell.casefold() for cell in rows[0])
                    groups.setdefault(key, []).append(name)
                if not found:
                    warnings.append(f'Trang {number}: chưa tách được bảng giá; kiểm tra PDF hoặc chuyển bảng thành XLSX trước khi nạp.')
                page.close()
    except PDFPasswordIncorrect:
        fail('PDF có mật khẩu. Mở khóa và lưu bản PDF trước khi nạp.')
    except HTTPException:
        raise
    except Exception:
        fail('Không đọc được PDF. File cần là PDF hợp lệ, không có mật khẩu.')
    if not sheets:
        fail('PDF chưa có bảng chữ đọc được. PDF ảnh scan cần OCR rồi kiểm tra/chuyển thành XLSX; chưa tạo giá từ ảnh.')
    # Only identical complete header rows permit a multi-page view; originals remain available.
    merged = {}
    for names in groups.values():
        if len(names) < 2:
            continue
        rows, refs = list(sheets[names[0]]), list(locations[names[0]])
        for name in names[1:]:
            rows.extend(sheets[name][1:])
            refs.extend(locations[name][1:])
        if len(rows) <= 10000:
            name = f'Gop cung tieu de {len(merged)+1}'
            merged[name] = rows
            locations[name] = refs
    if merged:
        warnings.append('Có bảng gộp cùng tiêu đề qua nhiều trang. Kiểm tra nguồn và điều kiện giá từng trang trước khi lưu; bảng từng trang vẫn giữ riêng.')
    sheets = dict(merged, **sheets)
    return sheets, dict(type='pdf', pages=pages, tables=count, warnings=warnings, locations=locations)


def pdf_price(value, row, role, number_format):
    if number_format not in ('vi', 'en', 'plain'):
        fail('Chọn định dạng giá trong PDF: dấu chấm hàng nghìn, dấu phẩy hàng nghìn hoặc số không phân cách.')
    if value is None or isinstance(value, str) and value.strip() in ('', '-', '—', '–'):
        return None
    if type(value) in (int, float):
        return value
    text = clean(value)
    text = re.sub(r'\s*(?:VNĐ|VND|đ|₫)\s*$', '', text, flags=re.IGNORECASE).strip()
    # Space groups are accepted only in the usual complete thousands pattern.
    if re.fullmatch(r'\d{1,3}(?: \d{3})+(?:[,.]\d+)?', text):
        text = text.replace(' ', '')
    pattern = {'vi': r'(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?',
               'en': r'(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?',
               'plain': r'\d+(?:\.\d+)?'}[number_format]
    if not re.fullmatch(pattern, text):
        fail(f'Dòng {row}, {role}: giá PDF “{text[:80]}” không đúng định dạng đã chọn. Kiểm tra giá và dấu phân cách; không tự suy ra công thức.')
    if number_format == 'vi':
        text = text.replace('.', '').replace(',', '.')
    elif number_format == 'en':
        text = text.replace(',', '')
    return float(text)


def excel_bytes(sheets, locations, filename, digest):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    book = Workbook()
    book.remove(book.active)
    for name, rows in sheets.items():
        sheet = book.create_sheet(name[:31])
        sheet.freeze_panes = 'A2'
        width = max(len(row) for row in rows)
        for index, row in enumerate(rows):
            ref = locations[name][index]
            source = f'{filename} | trang {ref["page"]}, bảng {ref["table"]}, dòng {ref["row"]} | SHA256 {digest}'
            values = list(row) + [''] * (width-len(row)) + ['Nguồn PDF' if index == 0 else source]
            for column, value in enumerate(values, 1):
                cell = sheet.cell(index+1, column, value)
                if isinstance(value, str):
                    cell.data_type = 's'  # PDF strings starting with '=' stay literal evidence.
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                if index == 0:
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='123F59')
        for column in range(1, width+2):
            sheet.column_dimensions[get_column_letter(column)].width = 35 if column in (1, width+1) else 22
        sheet.auto_filter.ref = sheet.dimensions
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()
