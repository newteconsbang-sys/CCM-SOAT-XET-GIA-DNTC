"""Thư viện giá: đọc Excel (pandas), kiểm tra, phân loại Vật tư / Nhân công / Ca máy,
quản lý version, sửa trực tiếp có lịch sử và so sánh hai version.

Không gọi dịch vụ AI bên ngoài: mọi nhận diện dựa trên tiêu đề cột, tên sheet, mã và ĐVT."""
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from backend.mep_db import get_connection, now_iso
from backend.mep_engine import normalize_mep_description
from backend.xlsx_utils import parse_number

ITEM_TYPES = ["Vật tư", "Nhân công", "Ca máy"]
MONEY_FIELDS = ("material_price", "labor_price", "machine_price", "total_price")
EDITABLE_FIELDS = {
    "item_type", "code", "description", "specification", "brand", "unit", "category",
    "material_price", "labor_price", "machine_price", "waste_percent", "total_price",
}
FIELD_LABELS = {
    "item_type": "Loại", "code": "Mã", "description": "Tên", "specification": "Quy cách", "brand": "Hãng",
    "unit": "ĐVT", "category": "Nhóm", "material_price": "Đơn giá VT", "labor_price": "Đơn giá NC",
    "machine_price": "Đơn giá máy", "waste_percent": "Hao hụt", "total_price": "Đơn giá tổng",
}

# Từ khóa tiêu đề cột (so sánh không dấu, chữ thường). Thứ tự quan trọng: cột cụ thể trước cột chung.
COLUMN_KEYS: List[Tuple[str, List[str]]] = [
    ("item_type", ["loai", "loai vt", "loai chi phi", "phan loai chi phi", "type"]),
    ("code", ["ma", "ma vt", "ma hieu", "ma vat tu", "ma so", "code"]),
    ("description", ["ten vat tu", "ten", "mo ta", "dien giai", "hang muc", "noi dung", "ten cong viec", "cong viec",
                     "ten may", "ten thiet bi", "ten vat lieu", "ten hang", "description",
                     # báo giá NCC (song ngữ): "Specifications / Đặc tính kỹ thuật", "Description of goods", "Item"
                     "specifications", "dac tinh ky thuat", "description of goods", "items", "item", "ten hang hoa", "hang hoa"]),
    ("specification", ["quy cach / hang", "quy cach", "thong so", "dac tinh", "chung loai", "specification"]),
    ("brand", ["hang", "nhan hieu", "thuong hieu", "hang sx", "brand"]),
    ("unit", ["dvt", "don vi", "don vi tinh", "dv", "d.v", "d/v", "unit", "uom"]),
    ("material_price", ["don gia vt", "don gia vat tu", "gia vt", "gia vat tu", "vat tu"]),
    ("labor_price", ["don gia nc", "don gia nhan cong", "gia nc", "gia nhan cong", "nhan cong"]),
    ("machine_price", ["don gia may", "don gia ca may", "gia may", "ca may", "may thi cong", "khau hao", "don gia khau hao"]),
    ("waste_percent", ["hao hut", "% hao hut", "ty le hao hut"]),
    ("total_price", ["don gia tong", "don gia tong hop", "don gia", "gia tong", "thanh gia", "don gia (vnd)", "gia",
                     "unit price", "unit rate", "price", "don gia ban"]),
    ("category", ["nhom", "nhom vat tu", "phan loai", "category"]),
    ("system", ["he", "he mep", "phan he", "system"]),
]


def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s or "")).replace("đ", "d").replace("Đ", "D")
    return re.sub(r"[̀-ͯ]", "", s).lower().strip()


def _clean(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def cell_json(v: Any) -> Any:
    """Giá trị ô → JSON (giữ số là số, ngày → chuỗi dd/mm/yyyy, trống → None)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return float(v) if isinstance(v, float) else int(v)
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%d/%m/%Y")
        except ValueError:
            return str(v)
    t = str(v).strip()
    return t or None


def _header_match(text: str, kws: List[str]) -> bool:
    t = re.sub(r"\s+", " ", strip_accents(text)).strip(" :*")
    if not t:
        return False
    return any(t == kw or t.startswith(kw + " ") or t.startswith(kw + "(") or t.startswith(kw + "/") for kw in kws)


def _map_heads(heads: Dict[int, str]) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    used = set()
    for key, kws in COLUMN_KEYS:
        for c, h in heads.items():
            if c in used or not h:
                continue
            if _header_match(h, kws):
                mapping[key] = c
                used.add(c)
                break
    return mapping


def detect_header(df: pd.DataFrame, max_scan: int = 15) -> Tuple[Optional[int], Dict[str, int], Dict[int, str]]:
    """Tìm dòng tiêu đề trong 15 dòng đầu → (dòng tiêu đề cuối 0-based, {trường: cột 0-based}, {cột: chữ tiêu đề}).
    Hỗ trợ tiêu đề 2 tầng như file Đơn giá ĐNTC: "ĐƠN GIÁ (VNĐ)" gộp ô, dòng dưới "VẬT TƯ" | "NHÂN CÔNG"."""
    best: Tuple[int, Optional[int], Dict[str, int], Dict[int, str]] = (0, None, {}, {})
    n = min(len(df), max_scan)
    for r in range(n):
        heads = {c: _clean(df.iat[r, c]) for c in range(df.shape[1])}
        cands = [(r, heads)]
        if r + 1 < len(df):
            sub = {c: _clean(df.iat[r + 1, c]) for c in range(df.shape[1])}
            # dòng dưới là tiêu đề con: có chữ, không có số
            if any(sub.values()) and not any(parse_number(df.iat[r + 1, c]) is not None and not isinstance(df.iat[r + 1, c], str)
                                             for c in range(df.shape[1])):
                cands.append((r + 1, {c: (sub[c] or heads[c]) for c in heads}))
        base_n = None
        for hr, hs in cands:
            mapping = _map_heads(hs)
            if hr != r and (base_n is None or len(mapping) <= base_n):
                continue  # chỉ dùng tiêu đề 2 tầng khi nó nhận thêm được cột (vd tách Vật tư / Nhân công)
            base_n = len(mapping) if hr == r else base_n
            score = len(mapping) + (3 if "description" in mapping else 0)
            if score > best[0]:
                best = (score, hr, mapping, hs)
    return best[1], best[2], best[3]


def detect_columns(df: pd.DataFrame, max_scan: int = 15) -> Tuple[Optional[int], Dict[str, int]]:
    hr, mapping, _ = detect_header(df, max_scan)
    return hr, mapping


def guess_item_type(raw_type: str, sheet: str, code: str, unit: str, desc: str) -> str:
    for src in (raw_type, sheet):
        t = strip_accents(src)
        if not t:
            continue
        if "nhan cong" in t or t in {"nc", "lao dong"}:
            return "Nhân công"
        if "may" in t or t in {"m", "cm"}:
            return "Ca máy"
        if "vat tu" in t or "vat lieu" in t or t in {"vt", "vl"}:
            return "Vật tư"
    c, u, d = strip_accents(code), strip_accents(unit), strip_accents(desc)
    if re.match(r"^nc[-_.\d]", c) or u in {"cong", "ngay cong", "gio cong"} or d.startswith("nhan cong"):
        return "Nhân công"
    if re.match(r"^(cm|m)[-_.]\d", c) or u in {"ca", "ca may", "gio may"} or d.startswith(("may ", "xe ", "can truc")):
        return "Ca máy"
    return "Vật tư"


def guess_category(desc: str, category: str = "", item_type: str = "Vật tư") -> str:
    if category:
        return category
    if item_type != "Vật tư":
        return item_type
    d = f" {normalize_mep_description(desc)} "
    rules = [("PCCC", ["sprinkler", "pccc", "chua chay", "bao chay", "binh cuu hoa"]),
             ("Cấp thoát nước", ["ppr", "upvc", "hdpe", "cap nuoc", "thoat nuoc", " van ", " bom ", "nuoc"]),
             ("Đèn", [" den ", " led ", "chieu sang"]),
             ("Ống & phụ kiện", [" ong ", " co ", " te ", "mang cap", "thang cap", "hop noi", "khop noi"]),
             ("Cáp điện", [" cap ", "day dien", " cv ", " cxv ", " cvv "]),
             ("Thiết bị điện", ["o cam", "cong tac", "aptomat", " mcb ", " mccb ", "tu dien", " rcbo "])]
    for name, kws in rules:
        if any(k in d for k in kws):
            return name
    return "Khác"


def item_key(it: Dict[str, Any]) -> str:
    """Khóa so sánh giữa hai version: Loại + Mã; nếu không có mã thì Loại + tên + ĐVT."""
    t = strip_accents(it.get("item_type") or "Vật tư")
    code = strip_accents(it.get("code") or "")
    if code:
        return f"{t}|{code}"
    return f"{t}|{normalize_mep_description(it.get('description') or '')}|{strip_accents(it.get('unit') or '')}"


def compute_total(material: Optional[float], labor: Optional[float], machine: Optional[float],
                  total: Optional[float]) -> Optional[float]:
    parts = [p for p in (material, labor, machine) if p is not None]
    if parts:
        return float(sum(parts))
    return total


# ------------------------------------------------------------------ đọc file

def _apply_percent_display(path: Path, structures: List[Dict[str, Any]]):
    """pandas không giữ định dạng ô: đọc lại định dạng bằng openpyxl để ô % hiển thị đúng như file (5%, 12,5%)."""
    if not structures:
        return
    import openpyxl
    from backend.xlsx_utils import percent_text
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return
    try:
        for st in structures:
            if st["sheet"] not in wb.sheetnames:
                continue
            ws = wb[st["sheet"]]
            pct: Dict[tuple, Any] = {}
            for row in ws.iter_rows(min_row=st["header_row"] + 1, min_col=st["first_col"], max_col=st["last_col"]):
                for c in row:
                    if getattr(c, "value", None) is not None and isinstance(c.value, (int, float)) and "%" in (c.number_format or ""):
                        pct[(c.row, c.column)] = percent_text(c.value)
            if not pct:
                continue
            for r in st["rows"]:
                for i in range(len(r["cells"])):
                    key = (r["row"], st["first_col"] + i)
                    if key in pct:
                        r["cells"][i] = pct[key]
    finally:
        wb.close()


def _hidden_sheets(path: Path) -> set:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        try:
            return {ws.title for ws in wb.worksheets if getattr(ws, "sheet_state", "visible") != "visible"}
        finally:
            wb.close()
    except Exception:
        return set()


def _raw_structure(sheet_name: str, df: pd.DataFrame) -> Dict[str, Any]:
    """Giữ nguyên sheet như file: dòng tiêu đề = dòng có nhiều chữ nhất trong 15 dòng đầu; mọi dòng sau giữ đúng thứ tự."""
    used_cols = [c for c in range(df.shape[1]) if df.iloc[:, c].notna().any()]
    first_c, last_c = (min(used_cols), max(used_cols)) if used_cols else (0, 0)
    best_r, best_n = 0, -1
    for r in range(min(len(df), 15)):
        n = sum(1 for c in range(first_c, last_c + 1) if isinstance(df.iat[r, c], str) and df.iat[r, c].strip())
        if n > best_n:
            best_r, best_n = r, n
    rows = []
    for r in range(0, best_r):  # dòng tiêu đề phía trên bảng (Hạng mục, Phương án…) → hiện làm chú thích trên bảng
        txt = " ".join(str(v).strip() for v in df.iloc[r, first_c:last_c + 1] if v is not None and str(v).strip() and str(v) != "nan")
        if txt:
            rows.append({"row": r + 1, "kind": "title", "cells": [txt]})
    for r in range(best_r + 1, len(df)):
        cells = [cell_json(df.iat[r, c]) for c in range(first_c, last_c + 1)]
        if any(v is not None for v in cells):
            rows.append({"row": r + 1, "kind": "raw", "cells": cells})
    return {"sheet": sheet_name, "header_row": best_r + 1, "first_col": first_c + 1, "last_col": last_c + 1,
            "headers": [_clean(df.iat[best_r, c]) for c in range(first_c, last_c + 1)], "mapping": {}, "rows": rows}


def parse_library_file(path: Path, raw_only: bool = False) -> Dict[str, Any]:
    """Đọc mọi sheet có cột Tên + Đơn giá. Trả entries (kèm errors/warnings từng dòng) và thống kê."""
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object, engine="openpyxl")
    except Exception as e:  # file hỏng, sai định dạng
        raise ValueError(f"Không đọc được file Excel: {e}") from e
    entries: List[Dict[str, Any]] = []
    errors: List[Dict[str, Any]] = []
    warnings: List[Dict[str, Any]] = []
    used_sheets: List[Dict[str, Any]] = []
    structures: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    seen: Dict[str, int] = {}
    hidden = _hidden_sheets(path)
    for sheet_name, df in sheets.items():
        if df.empty or df.dropna(how="all").empty or sheet_name in hidden:
            continue
        hr, cols, heads = detect_header(df)
        has_price = bool({"total_price", "material_price", "labor_price", "machine_price"} & cols.keys())
        if raw_only or hr is None or "description" not in cols or not has_price:
            # Sheet không phải bảng giá (vd chi tiết trực, kế hoạch vật tư): giữ nguyên để xem, không lấy giá
            structures.append(_raw_structure(sheet_name, df))
            continue
        head_cells = [c for c in range(df.shape[1]) if heads.get(c)]
        first_c = min(head_cells + list(cols.values()))
        last_c = max(head_cells + list(cols.values()))
        struct = {"sheet": sheet_name, "header_row": hr + 1, "first_col": first_c + 1, "last_col": last_c + 1,
                  "headers": [heads.get(c, "") for c in range(first_c, last_c + 1)],
                  "mapping": {k: int(v) + 1 for k, v in cols.items()}, "rows": []}
        used_sheets.append({"sheet": sheet_name, "header_row": hr + 1,
                            "columns": {k: int(v) + 1 for k, v in cols.items()}})
        structures.append(struct)
        for r in range(hr + 1, len(df)):
            row = df.iloc[r]
            cells = [cell_json(row.iloc[c]) for c in range(first_c, last_c + 1)]
            if all(v is None for v in cells):
                continue
            # Báo giá: chỉ lấy tới dòng "Cộng trước thuế / Sub-total"; bỏ VAT, tổng sau thuế, điều kiện, chữ ký phía dưới
            # (chỉ xét dòng không có ĐVT – dòng tổng, không phải hạng mục)
            no_unit = "unit" not in cols or not _clean(row.iloc[cols["unit"]])
            rtxt = strip_accents(" ".join(str(v) for v in cells if isinstance(v, str))).strip()
            if no_unit and re.search(r"(cong|tong).{0,20}(truoc (thue|vat)|chua (thue|vat))|sub-?total|total before (tax|vat)", rtxt):
                struct["rows"].append({"row": r + 1, "kind": "group", "cells": cells})
                break
            if no_unit and (re.search(r"(cong|tong).{0,20}sau thue|total after (tax|vat)|grand total", rtxt) or re.fullmatch(r"vat( \d.*)?", rtxt)):
                break
            get = lambda k: row.iloc[cols[k]] if k in cols else None
            desc = _clean(get("description"))
            code = _clean(get("code"))
            unit = _clean(get("unit"))
            raw = {k: get(k) for k in ("material_price", "labor_price", "machine_price", "total_price", "waste_percent")}
            vals = {k: parse_number(v) for k, v in raw.items()}
            # Dòng nhóm / tiêu đề mục / tổng: không có ĐVT và không có giá → giữ để hiển thị, không phải lỗi.
            if (not desc and not code) or (not unit and all(vals[k] is None for k in MONEY_FIELDS)):
                struct["rows"].append({"row": r + 1, "kind": "group", "cells": cells})
                continue
            itype = guess_item_type(_clean(get("item_type")), sheet_name, code, unit, desc)
            mat, lab, mac, tot = vals["material_price"], vals["labor_price"], vals["machine_price"], vals["total_price"]
            # Bảng chỉ có một cột "Đơn giá": gán vào đúng thành phần theo loại.
            if tot is not None and mat is None and lab is None and mac is None:
                if itype == "Nhân công":
                    lab = tot
                elif itype == "Ca máy":
                    mac = tot
            waste = vals["waste_percent"]
            if waste is not None and 0 < waste < 1 and isinstance(raw["waste_percent"], (int, float)):
                waste = round(waste * 100, 4)  # ô định dạng % (0,03 = 3%)
            computed = compute_total(mat, lab, mac, None)
            total = tot if tot is not None else computed
            e = {"sheet": sheet_name, "row": r + 1, "item_type": itype, "code": code, "description": desc,
                 "specification": _clean(get("specification")), "brand": _clean(get("brand")), "unit": unit,
                 "material_price": mat, "labor_price": lab, "machine_price": mac, "waste_percent": waste,
                 "total_price": total, "system": _clean(get("system")), "category": ""}
            e["category"] = guess_category(desc, _clean(get("category")), itype)
            problems, warns = [], []
            if not desc:
                problems.append("thiếu tên")
            if not unit:
                problems.append("thiếu ĐVT")
            for k in MONEY_FIELDS:
                if raw.get(k) not in (None, "") and _clean(raw.get(k)) and vals[k] is None:
                    problems.append(f"{FIELD_LABELS[k]} không phải số")
            if total is None or total <= 0:
                problems.append("thiếu đơn giá")
            if any(v is not None and v < 0 for v in (mat, lab, mac, total)):
                problems.append("đơn giá âm")
            k = item_key(e)
            if k in seen:
                problems.append(f"trùng với dòng {seen[k]}")
            else:
                seen[k] = r + 1
            if tot is not None and computed is not None and abs(computed - tot) > max(1.0, 0.01 * tot):
                warns.append(f"VT+NC+Máy ({computed:,.0f}) khác Đơn giá tổng ({tot:,.0f}) – giữ Đơn giá tổng của file")
            if waste is not None and not (0 <= waste <= 100):
                warns.append("hao hụt ngoài 0–100%")
            e["errors"], e["warnings"] = problems, warns
            if problems:
                errors.append({"sheet": sheet_name, "row": r + 1, "code": code, "description": desc, "problems": problems})
            if warns:
                warnings.append({"sheet": sheet_name, "row": r + 1, "code": code, "description": desc, "problems": warns})
            struct["rows"].append({"row": r + 1, "kind": "error" if problems else "item", "cells": cells, "entry": len(entries)})
            entries.append(e)
    _apply_percent_display(path, structures)
    if not entries and not (raw_only and any(st["rows"] for st in structures)):
        raise ValueError("Không tìm thấy bảng có cột Tên vật tư và Đơn giá trong file. Nếu là bảng Khấu hao / Nhân công trực, "
                         "chọn đúng Phân loại trước khi nạp.")
    by_type: Dict[str, int] = {}
    for e in entries:
        if not e["errors"]:
            by_type[e["item_type"]] = by_type.get(e["item_type"], 0) + 1
    return {"entries": entries, "errors": errors, "warnings": warnings, "sheets": used_sheets, "structure": structures, "skipped_sheets": skipped,
            "row_count": len(entries), "valid_count": len(entries) - len(errors), "by_type": by_type}


# ------------------------------------------------------------------ version

def ensure_library_schema():
    with get_connection() as conn:
        cols = {r[1] for r in conn.execute("PRAGMA table_info(mep_libraries)")}
        for name, typ in (("category", "TEXT"), ("waste_percent", "REAL"), ("item_type", "TEXT"),
                          ("machine_price", "REAL"), ("updated_at", "TEXT")):
            if name not in cols:
                conn.execute(f"ALTER TABLE mep_libraries ADD COLUMN {name} {typ}")
        conn.execute("UPDATE mep_libraries SET item_type='Vật tư' WHERE item_type IS NULL OR item_type=''")
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS mep_library_edits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version_id INTEGER NOT NULL,
            item_id INTEGER,
            action TEXT NOT NULL,
            field TEXT,
            old_value TEXT,
            new_value TEXT,
            item_label TEXT,
            changed_by TEXT,
            changed_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_lib_edits_ver ON mep_library_edits(version_id);
        CREATE INDEX IF NOT EXISTS idx_mep_lib_version ON mep_libraries(version);
        """)


def active_version_id(conn) -> Optional[int]:
    r = conn.execute("SELECT id FROM mep_library_versions WHERE status='active' ORDER BY id DESC LIMIT 1").fetchone()
    return r[0] if r else None


def next_version_name(conn, today: Optional[str] = None) -> str:
    today = today or now_iso()[:10]
    base = "v" + today[:7].replace("-", ".")
    names = {r[0] for r in conn.execute("SELECT name FROM mep_library_versions")}
    if base not in names:
        return base
    i = 2
    while f"{base}-{i}" in names:
        i += 1
    return f"{base}-{i}"


def _row_values(e: Dict[str, Any], vid: int, version_name: str, date_val: str) -> tuple:
    system = e.get("system") or {"Cấp thoát nước": "Cấp thoát nước", "PCCC": "PCCC"}.get(e.get("category"), "Điện")
    total = compute_total(e.get("material_price"), e.get("labor_price"), e.get("machine_price"), e.get("total_price")) \
        if e.get("total_price") is None else e.get("total_price")
    return (system, e.get("code", ""), e["description"], e.get("specification", ""), e.get("brand", ""), "", "",
            e.get("unit", ""), e.get("material_price"), e.get("labor_price"), float(total or 0), f"Version {version_name}",
            "", "", date_val, vid,
            normalize_mep_description(f"{e['description']} {e.get('specification', '')} {e.get('brand', '')}"),
            now_iso(), e.get("category") or guess_category(e["description"], "", e.get("item_type") or "Vật tư"),
            e.get("waste_percent"), e.get("item_type") or "Vật tư", e.get("machine_price"))


INSERT_SQL = """INSERT INTO mep_libraries (system, code, description, specification, brand, origin, model, unit,
    material_price, labor_price, total_price, source, vendor, project, date, version, normalized_desc,
    created_at, category, waste_percent, item_type, machine_price) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""


def create_version(name: str, entries: List[Dict[str, Any]], effective_date: str = "", note: str = "",
                   created_by: str = "", status: str = "checked", error_count: int = 0,
                   scope: str = "Toàn bộ (Điện + Nước)", structure: Optional[List[Dict[str, Any]]] = None,
                   source_bytes: Optional[bytes] = None, source_name: str = "") -> Dict[str, Any]:
    name = (name or "").strip()
    valid_idx = [i for i, e in enumerate(entries) if not e.get("errors") and (e.get("description") or "").strip()]
    valid = [entries[i] for i in valid_idx]
    # Thư viện chỉ để tra cứu (Khấu hao, Nhân công trực…): không có mục giá vẫn lưu được nếu có sheet giữ nguyên
    if not valid and not any(st.get("rows") for st in (structure or [])):
        raise ValueError("Không có dòng hợp lệ để tạo version")
    effective_date = effective_date or now_iso()[:10]
    with get_connection() as conn:
        if not name:
            name = next_version_name(conn, effective_date)
        if conn.execute("SELECT 1 FROM mep_library_versions WHERE name=?", (name,)).fetchone():
            raise ValueError(f"Version {name} đã tồn tại")
        cur = conn.execute(
            "INSERT INTO mep_library_versions (name, effective_date, scope, note, status, item_count, error_count, created_by, created_at)"
            " VALUES (?,?,?,?,?,?,?,?,?)",
            (name, effective_date, scope, note, status, len(valid), error_count, created_by, now_iso()))
        vid = cur.lastrowid
        item_ids: Dict[int, int] = {}
        for i in valid_idx:
            item_ids[i] = conn.execute(INSERT_SQL, _row_values(entries[i], vid, name, effective_date)).lastrowid
        if structure:
            from backend import library_structure
            library_structure.save_structure(conn, vid, structure, item_ids, source_bytes, source_name)
        conn.execute("INSERT INTO mep_library_edits (version_id, action, new_value, changed_by, changed_at) VALUES (?,?,?,?,?)",
                     (vid, "import", f"{len(valid)} mục" + (f" · bỏ {error_count} dòng lỗi" if error_count else ""), created_by, now_iso()))
    return {"id": vid, "name": name, "count": len(valid)}


def activate_version(vid: int) -> None:
    with get_connection() as conn:
        if not conn.execute("SELECT 1 FROM mep_library_versions WHERE id=?", (vid,)).fetchone():
            raise LookupError("Không tìm thấy version")
        conn.execute("UPDATE mep_library_versions SET status='archived' WHERE status='active' AND id<>?", (vid,))
        conn.execute("UPDATE mep_library_versions SET status='active' WHERE id=?", (vid,))


def _refresh_count(conn, vid: int):
    n = conn.execute("SELECT COUNT(*) FROM mep_libraries WHERE version=?", (vid,)).fetchone()[0]
    conn.execute("UPDATE mep_library_versions SET item_count=? WHERE id=?", (n, vid))


def version_rows(conn, vid: int) -> List[Dict[str, Any]]:
    return [dict(r) for r in conn.execute(
        "SELECT * FROM mep_libraries WHERE version=? ORDER BY item_type, code, description", (vid,))]


def previous_version_id(conn, vid: int) -> Optional[int]:
    r = conn.execute(
        """SELECT id FROM mep_library_versions WHERE id<>? AND status<>'reference' AND (COALESCE(effective_date,''), id) <
               (SELECT COALESCE(effective_date,''), id FROM mep_library_versions WHERE id=?)
           ORDER BY COALESCE(effective_date,'') DESC, id DESC LIMIT 1""", (vid, vid)).fetchone()
    return r[0] if r else None


def list_items(vid: int, item_type: str = "", category: str = "", q: str = "",
               compare_to: Optional[int] = None, sheet_id: Optional[int] = None) -> Dict[str, Any]:
    with get_connection() as conn:
        if not conn.execute("SELECT 1 FROM mep_library_versions WHERE id=?", (vid,)).fetchone():
            raise LookupError("Không tìm thấy version")
        sql, params = "SELECT * FROM mep_libraries WHERE version=?", [vid]
        if item_type:
            sql += " AND COALESCE(item_type,'Vật tư')=?"; params.append(item_type)
        if category:
            sql += " AND COALESCE(category, system)=?"; params.append(category)
        rows = [dict(r) for r in conn.execute(sql + " ORDER BY code, description", params)]
        if q.strip():
            nq = strip_accents(q)
            rows = [r for r in rows if nq in strip_accents(f"{r['code']} {r['description']} {r['specification']} {r['brand']}")]
        if compare_to is None:
            compare_to = previous_version_id(conn, vid)
        base = {item_key(r): r for r in version_rows(conn, compare_to)} if compare_to else {}
        for r in rows:
            r["item_type"] = r.get("item_type") or "Vật tư"
            old = base.get(item_key(r))
            r["compare_percent"] = round((r["total_price"] - old["total_price"]) / old["total_price"] * 100, 1) \
                if old and old["total_price"] else None
            r["is_new"] = bool(base) and old is None
        types = conn.execute("SELECT COALESCE(item_type,'Vật tư') t, COUNT(*) n FROM mep_libraries WHERE version=? GROUP BY t", (vid,)).fetchall()
        cat_sql, cat_params = "SELECT COALESCE(category, system) c, COUNT(*) n FROM mep_libraries WHERE version=?", [vid]
        if item_type:
            cat_sql += " AND COALESCE(item_type,'Vật tư')=?"; cat_params.append(item_type)
        cats = conn.execute(cat_sql + " GROUP BY c ORDER BY c", cat_params).fetchall()
        cv = conn.execute("SELECT name FROM mep_library_versions WHERE id=?", (compare_to,)).fetchone() if compare_to else None
        tcount = {t["t"]: t["n"] for t in types}
        from backend import library_structure
        sheet_list = library_structure.sheets(conn, vid)
        sheet_rows = None
        if sheet_list:
            sid = sheet_id if any(x["id"] == sheet_id for x in sheet_list) else sheet_list[0]["id"]
            sheet_rows = {"sheet_id": sid, "rows": library_structure.sheet_rows(
                conn, sid, {r["id"]: r for r in rows}, bool(item_type or category or q.strip()))}
        return {"items": rows, "sheets": sheet_list, "sheet": sheet_rows,
                "types": [{"name": t, "count": tcount.get(t, 0)} for t in ITEM_TYPES],
                "categories": [{"name": c["c"] or "Khác", "count": c["n"]} for c in cats],
                "compare_to": compare_to, "compare_name": cv["name"] if cv else None}


def _validate_patch(patch: Dict[str, Any]) -> Dict[str, Any]:
    clean: Dict[str, Any] = {}
    for k, v in patch.items():
        if k not in EDITABLE_FIELDS:
            continue
        if k in MONEY_FIELDS or k == "waste_percent":
            if v in (None, ""):
                clean[k] = None
                continue
            n = parse_number(v)
            if n is None:
                raise ValueError(f"{FIELD_LABELS[k]} phải là số")
            if n < 0:
                raise ValueError(f"{FIELD_LABELS[k]} không được âm")
            if k == "waste_percent" and n > 100:
                raise ValueError("Hao hụt tối đa 100%")
            clean[k] = n
        elif k == "item_type":
            if v not in ITEM_TYPES:
                raise ValueError("Loại phải là Vật tư, Nhân công hoặc Ca máy")
            clean[k] = v
        else:
            clean[k] = str(v or "").strip()
    if "description" in clean and not clean["description"]:
        raise ValueError("Tên không được để trống")
    if "unit" in clean and not clean["unit"]:
        raise ValueError("ĐVT không được để trống")
    return clean


def update_item(item_id: int, patch: Dict[str, Any], user: str = "") -> Dict[str, Any]:
    clean = _validate_patch(patch)
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM mep_libraries WHERE id=?", (item_id,)).fetchone()
        if not row:
            raise LookupError("Không tìm thấy mục thư viện")
        cur = dict(row)
        new = {**cur, **clean}
        # Sửa thành phần giá thì tự tính lại Đơn giá tổng (trừ khi người dùng sửa thẳng Đơn giá tổng).
        if ({"material_price", "labor_price", "machine_price"} & clean.keys()) and "total_price" not in clean:
            t = compute_total(new.get("material_price"), new.get("labor_price"), new.get("machine_price"), None)
            if t is not None:
                new["total_price"] = t
        if not new.get("total_price") or new["total_price"] <= 0:
            raise ValueError("Đơn giá tổng phải lớn hơn 0")
        changed = {k: new[k] for k in EDITABLE_FIELDS if new.get(k) != cur.get(k)}
        if not changed:
            return {"updated": False, "item": cur}
        if {"description", "specification", "brand"} & changed.keys():
            changed["normalized_desc"] = normalize_mep_description(f"{new['description']} {new.get('specification') or ''} {new.get('brand') or ''}")
        conn.execute(f"UPDATE mep_libraries SET {', '.join(k + '=?' for k in changed)}, updated_at=? WHERE id=?",
                     (*changed.values(), now_iso(), item_id))
        label = f"{cur.get('code') or ''} {cur['description']}".strip()
        for k, v in changed.items():
            if k == "normalized_desc":
                continue
            conn.execute("INSERT INTO mep_library_edits (version_id, item_id, action, field, old_value, new_value, item_label, changed_by, changed_at)"
                         " VALUES (?,?,?,?,?,?,?,?,?)",
                         (cur["version"], item_id, "edit", k, _txt(cur.get(k)), _txt(v), label, user, now_iso()))
            if k == "total_price":  # bảng lịch sử giá cũ vẫn được ghi để không đứt mạch dữ liệu
                conn.execute("INSERT INTO mep_price_history (library_id, old_price, new_price, changed_by, reason, timestamp) VALUES (?,?,?,?,?,?)",
                             (item_id, cur.get("total_price"), v, user or "web", "Sửa trực tiếp trên Thư viện giá", now_iso()))
        out = dict(conn.execute("SELECT * FROM mep_libraries WHERE id=?", (item_id,)).fetchone())
    return {"updated": True, "item": out, "changed": [k for k in changed if k != "normalized_desc"]}


def add_item(vid: int, data: Dict[str, Any], user: str = "", sheet_id: Optional[int] = None) -> Dict[str, Any]:
    clean = _validate_patch(data)
    if not clean.get("description") or not clean.get("unit"):
        raise ValueError("Nhập Tên và ĐVT")
    clean.setdefault("item_type", "Vật tư")
    clean["total_price"] = clean.get("total_price") if clean.get("total_price") is not None else \
        compute_total(clean.get("material_price"), clean.get("labor_price"), clean.get("machine_price"), None)
    if not clean["total_price"] or clean["total_price"] <= 0:
        raise ValueError("Nhập đơn giá lớn hơn 0")
    with get_connection() as conn:
        v = conn.execute("SELECT name, effective_date FROM mep_library_versions WHERE id=?", (vid,)).fetchone()
        if not v:
            raise LookupError("Không tìm thấy version")
        existing = {item_key(r) for r in version_rows(conn, vid)}
        if item_key(clean) in existing:
            raise ValueError("Mã/tên này đã có trong version")
        cur = conn.execute(INSERT_SQL, _row_values(clean, vid, v["name"], v["effective_date"] or now_iso()[:10]))
        iid = cur.lastrowid
        if sheet_id:
            from backend import library_structure
            library_structure.link_new_item(conn, vid, int(sheet_id), iid)
        conn.execute("INSERT INTO mep_library_edits (version_id, item_id, action, new_value, item_label, changed_by, changed_at) VALUES (?,?,?,?,?,?,?)",
                     (vid, iid, "add", _txt(clean["total_price"]), f"{clean.get('code') or ''} {clean['description']}".strip(), user, now_iso()))
        _refresh_count(conn, vid)
        return dict(conn.execute("SELECT * FROM mep_libraries WHERE id=?", (iid,)).fetchone())


def delete_item(item_id: int, user: str = "") -> None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM mep_libraries WHERE id=?", (item_id,)).fetchone()
        if not row:
            raise LookupError("Không tìm thấy mục thư viện")
        conn.execute("DELETE FROM mep_libraries WHERE id=?", (item_id,))
        conn.execute("INSERT INTO mep_library_edits (version_id, item_id, action, old_value, item_label, changed_by, changed_at) VALUES (?,?,?,?,?,?,?)",
                     (row["version"], item_id, "delete", _txt(row["total_price"]), f"{row['code'] or ''} {row['description']}".strip(), user, now_iso()))
        _refresh_count(conn, row["version"])


def history(vid: int, limit: int = 300) -> List[Dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute("SELECT * FROM mep_library_edits WHERE version_id=? ORDER BY id DESC LIMIT ?", (vid, limit)).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d["field_label"] = FIELD_LABELS.get(d.get("field") or "", d.get("field") or "")
        out.append(d)
    return out


def _txt(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ------------------------------------------------------------------ so sánh

COMPARE_FIELDS = ["description", "specification", "unit", "material_price", "labor_price", "machine_price", "waste_percent", "total_price"]


def _same(a: Any, b: Any) -> bool:
    if isinstance(a, (int, float)) or isinstance(b, (int, float)):
        if a is None or b is None:
            return a is None and b is None
        return abs(float(a) - float(b)) < 0.005
    return strip_accents(a or "") == strip_accents(b or "")


def compare_versions(base_id: int, target_id: int) -> Dict[str, Any]:
    """So sánh version gốc (base) với version đích (target): thêm mới / thay đổi / bị xóa."""
    if base_id == target_id:
        raise ValueError("Chọn hai version khác nhau")
    with get_connection() as conn:
        names = {r["id"]: r["name"] for r in conn.execute(
            "SELECT id, name FROM mep_library_versions WHERE id IN (?,?)", (base_id, target_id))}
        if len(names) < 2:
            raise LookupError("Không tìm thấy version cần so sánh")
        a_rows = version_rows(conn, base_id)
        b_rows = version_rows(conn, target_id)
    a: Dict[str, Dict[str, Any]] = {}
    for r in a_rows:
        a.setdefault(item_key(r), r)
    bm: Dict[str, Dict[str, Any]] = {}
    for r in b_rows:
        bm.setdefault(item_key(r), r)
    added_keys = [k for k in bm if k not in a]
    removed_keys = [k for k in a if k not in bm]
    common = [k for k in bm if k in a]

    def brief(r: Dict[str, Any]) -> Dict[str, Any]:
        d = {k: r.get(k) for k in ["item_type", "code", "description", "specification", "unit", "material_price",
                                   "labor_price", "machine_price", "waste_percent", "total_price"]}
        d["item_type"] = d["item_type"] or "Vật tư"
        return d

    changed = []
    for k in common:
        ra, rb = brief(a[k]), brief(bm[k])
        diffs = [f for f in COMPARE_FIELDS if not _same(ra.get(f), rb.get(f))]
        if diffs:
            old_t, new_t = ra.get("total_price") or 0, rb.get("total_price") or 0
            changed.append({**rb, "old": {f: ra.get(f) for f in diffs}, "fields": diffs,
                            "field_labels": [FIELD_LABELS[f] for f in diffs],
                            "old_total": old_t, "diff_percent": round((new_t - old_t) / old_t * 100, 2) if old_t else None})
    changed.sort(key=lambda c: -abs(c["diff_percent"] or 0))
    return {"base": {"id": base_id, "name": names[base_id]}, "target": {"id": target_id, "name": names[target_id]},
            "added": [brief(bm[k]) for k in added_keys],
            "removed": [brief(a[k]) for k in removed_keys],
            "changed": changed, "unchanged": len(common) - len(changed),
            "summary": {"added": len(added_keys), "removed": len(removed_keys), "changed": len(changed),
                        "unchanged": len(common) - len(changed)}}


def preview_against(entries: List[Dict[str, Any]], base_id: Optional[int]) -> Dict[str, Any]:
    """Thống kê nhanh file mới so với version đang áp dụng (trước khi lưu)."""
    if not base_id:
        return {"changed": [], "new": [e for e in entries if not e.get("errors")], "removed": []}
    with get_connection() as conn:
        base = {item_key(r): r for r in version_rows(conn, base_id)}
    seen, changed, new = set(), [], []
    for e in entries:
        if e.get("errors"):
            continue
        k = item_key(e)
        seen.add(k)
        old = base.get(k)
        if not old:
            new.append({"code": e["code"], "description": e["description"], "new": e["total_price"]})
        elif not _same(old["total_price"], e["total_price"]):
            changed.append({"code": e["code"], "description": e["description"], "old": old["total_price"], "new": e["total_price"]})
    removed = [{"code": r["code"], "description": r["description"], "old": r["total_price"]} for k, r in base.items() if k not in seen]
    return {"changed": changed, "new": new, "removed": removed}
