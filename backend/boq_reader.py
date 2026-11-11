"""Nhận diện bảng BOQ / báo giá NTP trong Excel: dòng tiêu đề (kể cả tiêu đề 2 tầng có ô gộp)
và vị trí các cột STT, Hạng mục, ĐVT, Khối lượng, Đơn giá VT/NC/Khấu hao, Đơn giá tổng, Thành tiền, Ghi chú."""
import re
from typing import Any, Dict, List, Optional, Tuple

from backend.library_service import strip_accents

FIELDS: List[Tuple[str, List[str]]] = [
    ("stt", ["stt", "tt", "so tt", "no", "no.", "so thu tu"]),
    ("description", ["hang muc", "noi dung", "noi dung cong viec", "mo ta", "dien giai", "ten vat tu",
                     "ten cong viec", "ten hang muc", "ten", "description", "ten vat tu thiet bi"]),
    ("code", ["ma hieu", "ma vt", "ma so", "ma cong viec", "ma hieu don gia", "ma", "code", "item code"]),
    ("specification", ["thong so", "thong so ky thuat", "quy cach", "dac tinh", "tieu chuan", "specification"]),
    ("brand", ["nhan hieu", "thuong hieu", "hang sx", "hang san xuat", "xuat xu", "hang", "brand"]),
    ("model", ["model", "ky hieu"]),
    ("unit", ["dvt", "don vi", "don vi tinh", "unit"]),
    ("quantity", ["khoi luong", "kl", "so luong", "sl", "qty", "quantity", "khoi luong hop dong"]),
    ("material_price", ["don gia vat tu", "don gia vt", "don gia vl", "don gia vat lieu", "gia vat tu", "gia vt",
                        "vat tu", "vat lieu", "vl", "vt"]),
    ("labor_price", ["don gia nhan cong", "don gia nc", "gia nhan cong", "gia nc", "nhan cong", "nc"]),
    ("dep_done_pct", []),  # "% ĐÃ KH": phần đã khấu hao trước đó – nhận diện riêng trong _score
    ("depreciation_pct", []),  # "% KH vật tư", "Tỷ lệ khấu hao" – nhận diện riêng trong _score
    ("start_date", ["thoi gian bat dau", "ngay bat dau", "bat dau", "tu ngay", "start date", "start"]),
    ("end_date", ["thoi gian ket thuc", "ngay ket thuc", "ket thuc", "den ngay", "end date", "finish"]),
    ("machine_price", ["don gia khau hao", "khau hao", "don gia may", "don gia ca may", "may thi cong", "ca may", "may", "m"]),
    ("price", ["don gia tong", "don gia tong hop", "tong don gia", "don gia", "don gia du thau", "unit price", "price", "don gia (vnd)"]),
    ("amount", ["thanh tien", "tong thanh tien", "thanh tien tong", "tong tien", "gia tri", "amount", "total"]),
    ("note", ["ghi chu", "note", "notes", "remark", "remarks"]),
]
SUB_HINT = {"vat tu", "vt", "vl", "vat lieu", "nhan cong", "nc", "may", "m", "khau hao", "ca may", "tong", "cong", "tong cong"}
PART_WORDS = ("vat tu", "vat lieu", "nhan cong", "khau hao", " may", " vt", " nc", " vl")


def _norm(v: Any) -> str:
    return re.sub(r"\s+", " ", strip_accents(str(v or ""))).strip(" :*.-")


def _row_texts(ws, r: int, max_col: int) -> Dict[int, str]:
    vals = {c: ws.cell(row=r, column=c).value for c in range(1, max_col + 1)}
    for mr in ws.merged_cells.ranges:  # ô gộp: lan giá trị ra mọi cột trong vùng
        if mr.min_row <= r <= mr.max_row:
            v = ws.cell(row=mr.min_row, column=mr.min_col).value
            for c in range(mr.min_col, min(mr.max_col, max_col) + 1):
                vals[c] = v
    return {c: str(v).strip() for c, v in vals.items() if v is not None and str(v).strip()}


def _score(field: str, text: str, kws: List[str]) -> int:
    t = _norm(text)
    if not t:
        return 0
    if field in ("depreciation_pct", "dep_done_pct"):
        raw = strip_accents(str(text))
        is_pct = "%" in raw or "ty le" in t or "he so" in t
        is_kh = bool(re.search(r"\bkh\b", t)) or "khau hao" in t
        done = bool(re.search(r"\bda (kh|khau hao)\b", t))
        if field == "dep_done_pct":
            return 6 if is_kh and done else 0
        return 5 if is_pct and is_kh and not done else 0
    best = 0
    for kw in kws:
        if t == kw:
            best = max(best, 4)
        elif t.startswith(kw + " ") or t.startswith(kw + "(") or t.startswith(kw + "/"):
            best = max(best, 2)
    if best and field in ("price", "amount"):
        padded = f" {t} "
        if any(p in padded for p in PART_WORDS):
            best -= 2  # "Thành tiền vật tư" không phải cột tổng
        elif "tong" in t or "cong" in t:
            best += 1
    return best


def map_columns(headers: Dict[int, str]) -> Dict[str, Optional[int]]:
    mapping: Dict[str, Optional[int]] = {f: None for f, _ in FIELDS}
    used = set()
    for field, kws in FIELDS:
        cands = [(s, -c, c) for c, h in headers.items() if c not in used and (s := _score(field, h, kws)) > 0]
        if cands:
            cands.sort(reverse=True)
            mapping[field] = cands[0][2]
            used.add(cands[0][2])
    return mapping


def detect_header(ws, scan_rows: int = 20, max_col: int = 40) -> Dict[str, Any]:
    """Trả {'header_row': dòng tiêu đề cuối (dữ liệu bắt đầu ngay sau), 'headers': {cột: tiêu đề}, 'mapping': ...}."""
    max_col = min(ws.max_column or 1, max_col)
    best: Tuple[int, int, Dict[int, str], Dict[str, Optional[int]]] = (0, 1, {}, {})
    last = min(ws.max_row or 1, scan_rows)
    for r in range(1, last + 1):
        top = _row_texts(ws, r, max_col)
        if not top:
            continue
        options = [(r, top)]
        if r + 1 <= (ws.max_row or 1):
            sub = _row_texts(ws, r + 1, max_col)
            if sub and sum(1 for v in sub.values() if _norm(v) in SUB_HINT) >= 2:
                combined = {c: f"{top.get(c, '')} {sub.get(c, '')}".strip() if sub.get(c) and _norm(sub.get(c)) != _norm(top.get(c, ''))
                            else top.get(c, sub.get(c, "")) for c in set(top) | set(sub)}
                options.append((r + 1, combined))
        for hr, heads in options:
            m = map_columns(heads)
            score = sum(1 for v in m.values() if v) + (4 if m.get("description") else 0) + \
                (3 if (m.get("price") or m.get("material_price")) else 0) + (1 if hr != r else 0)
            if score > best[0]:
                best = (score, hr, heads, m)
    return {"header_row": best[1], "headers": best[2], "mapping": best[3]}


def is_index_row(desc: str) -> bool:
    """Dòng đánh số cột kiểu (1) (2) (3) ngay dưới tiêu đề."""
    return bool(re.fullmatch(r"\(?\s*\d{1,2}\s*\)?", str(desc or "").strip()))
