"""Tiện ích dùng chung: đọc số kiểu Việt Nam, ghi bảng Excel, trả file tải về."""
import io
import math
import re
from typing import Any, List, Optional
from urllib.parse import quote

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from fastapi.responses import StreamingResponse

HEAD_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
HEAD_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

_CURRENCY = re.compile(r"(vnđ|vnd|đồng|đ|₫)", re.I)


def parse_number(value: Any) -> Optional[float]:
    """Đọc số từ ô Excel hoặc chuỗi: 1.250.000 / 1,250,000 / 25.000,5 / 18500.5 / '12%'.

    Trả None khi ô trống hoặc không phải số (không đoán)."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        f = float(value)
        return None if math.isnan(f) or math.isinf(f) else f
    s = _CURRENCY.sub("", str(value)).strip().replace(" ", "").replace(" ", "")
    if not s or s in {"-", "—", "–"}:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").rstrip("%")
    if not re.fullmatch(r"[-+]?[\d.,]+", s):
        return None
    if "." in s and "," in s:
        dec = "." if s.rfind(".") > s.rfind(",") else ","
        s = s.replace("," if dec == "." else ".", "").replace(dec, ".")
    elif s.count(".") > 1 or s.count(",") > 1:
        s = s.replace(".", "").replace(",", "")
    elif "," in s or "." in s:
        sep = "," if "," in s else "."
        tail = s.split(sep)[1]
        s = s.replace(sep, "") if len(tail) == 3 else s.replace(sep, ".")
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f


def xlsx_response(wb: openpyxl.Workbook, filename: str) -> StreamingResponse:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


def bytes_response(data: bytes, filename: str, media_type: str) -> StreamingResponse:
    return StreamingResponse(io.BytesIO(data), media_type=media_type,
                             headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"})


def write_table(ws, start_row: int, headers: List[str], rows: List[List[Any]], widths: List[int]):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.fill, cell.font, cell.border = HEAD_FILL, HEAD_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r, row in enumerate(rows, start_row + 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = "#,##0.##"
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def percent_text(value) -> str:
    """0.125 → '12,5%' (ô Excel định dạng phần trăm)."""
    p = round(float(value) * 100, 2)
    txt = f"{p:.2f}".rstrip("0").rstrip(".")
    return txt.replace(".", ",") + "%"


def display_cell(cell):
    """Giá trị hiển thị của ô openpyxl: ô định dạng % thì trả chuỗi '5%' như Excel hiện."""
    v = cell.value
    if isinstance(v, (int, float)) and not isinstance(v, bool) and "%" in (cell.number_format or ""):
        return percent_text(v)
    return v
