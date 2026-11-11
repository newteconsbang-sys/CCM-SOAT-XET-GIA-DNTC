"""Giữ nguyên cấu trúc file thư viện người dùng nạp: từng sheet, cột gốc (tiêu đề + thứ tự),
thứ tự dòng kể cả dòng nhóm/tổng. Xuất lại đúng file gốc, chỉ ghi các ô người dùng đã sửa."""
import io
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl

from backend.mep_db import get_connection, now_iso

ROOT = Path(__file__).resolve().parents[1]
FILES_DIR = ROOT / "data" / "library_files"
EDITABLE = ("code", "description", "specification", "brand", "unit", "material_price", "labor_price",
            "machine_price", "waste_percent", "total_price", "item_type", "category")


def ensure_schema():
    with get_connection() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS mep_library_sheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version_id INTEGER NOT NULL,
            sheet_index INTEGER,
            sheet_name TEXT NOT NULL,
            header_row INTEGER,
            first_col INTEGER,
            last_col INTEGER,
            headers_json TEXT,
            mapping_json TEXT
        );
        CREATE TABLE IF NOT EXISTS mep_library_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version_id INTEGER NOT NULL,
            sheet_id INTEGER NOT NULL,
            excel_row INTEGER,
            seq INTEGER,
            kind TEXT,
            cells_json TEXT,
            item_id INTEGER
        );
        CREATE INDEX IF NOT EXISTS idx_lib_rows_sheet ON mep_library_rows(sheet_id, seq);
        CREATE INDEX IF NOT EXISTS idx_lib_rows_item ON mep_library_rows(item_id);
        """)
        cols = {r[1] for r in conn.execute("PRAGMA table_info(mep_library_versions)")}
        if "source_file" not in cols:
            conn.execute("ALTER TABLE mep_library_versions ADD COLUMN source_file TEXT")
        # Phân loại thư viện: dntc = Đơn giá ĐNTC chuẩn, ncc = Báo giá NCC, hd = Hợp đồng tham khảo, kh = Khấu hao, truc = Đơn giá nhân công trực
        if "kind" not in cols:
            conn.execute("ALTER TABLE mep_library_versions ADD COLUMN kind TEXT DEFAULT 'dntc'")
        trim_ncc_footer(conn)


def trim_ncc_footer(conn) -> int:
    """Báo giá NCC đã nạp trước đây: chỉ giữ tới dòng "Cộng giá trước thuế / Sub-total",
    bỏ các dòng VAT, cộng sau thuế, điều khoản, chữ ký phía dưới (không đụng hạng mục giá)."""
    import re
    from backend.library_service import strip_accents
    pre = re.compile(r"(cong|tong).{0,20}(truoc (thue|vat)|chua (thue|vat))|sub-?total|total before (tax|vat)")
    post = re.compile(r"(cong|tong).{0,20}sau thue|total after (tax|vat)|grand total")
    removed = 0
    sheets = conn.execute("SELECT s.id, s.mapping_json, s.first_col FROM mep_library_sheets s JOIN mep_library_versions v ON v.id=s.version_id "
                          "WHERE COALESCE(v.kind,'dntc')='ncc' AND COALESCE(s.mapping_json,'') NOT IN ('','{}','null')").fetchall()
    for sh in sheets:
        try:
            mp = json.loads(sh["mapping_json"] or "{}")
        except Exception:
            continue
        ui = (mp["unit"] - (sh["first_col"] or 1)) if isinstance(mp.get("unit"), int) else None
        cut = None
        for r in conn.execute("SELECT seq, cells_json, item_id FROM mep_library_rows WHERE sheet_id=? ORDER BY seq", (sh["id"],)):
            if r["item_id"]:
                continue
            c = json.loads(r["cells_json"] or "[]")
            if ui is not None and 0 <= ui < len(c) and c[ui] not in (None, ""):
                continue
            t = re.sub(r"\s+", " ", strip_accents(" ".join(str(v) for v in c if isinstance(v, str)))).strip().lower()
            if pre.search(t):
                cut = (r["seq"], False)
                break
            if post.search(t) or re.fullmatch(r"vat( \d.*)?", t):
                cut = (r["seq"], True)
                break
        if cut:
            op = ">=" if cut[1] else ">"
            removed += conn.execute(f"DELETE FROM mep_library_rows WHERE sheet_id=? AND item_id IS NULL AND seq {op} ?",
                                    (sh["id"], cut[0])).rowcount
    return removed


def save_structure(conn, vid: int, structure: List[Dict[str, Any]], item_ids: Dict[int, int],
                   source_bytes: Optional[bytes], source_name: str = ""):
    for idx, st in enumerate(structure):
        sid = conn.execute(
            "INSERT INTO mep_library_sheets (version_id, sheet_index, sheet_name, header_row, first_col, last_col, headers_json, mapping_json)"
            " VALUES (?,?,?,?,?,?,?,?)",
            (vid, idx, st["sheet"], st["header_row"], st["first_col"], st["last_col"],
             json.dumps(st["headers"], ensure_ascii=False), json.dumps(st["mapping"]))).lastrowid
        conn.executemany(
            "INSERT INTO mep_library_rows (version_id, sheet_id, excel_row, seq, kind, cells_json, item_id) VALUES (?,?,?,?,?,?,?)",
            [(vid, sid, r["row"], n, r["kind"], json.dumps(r["cells"], ensure_ascii=False), item_ids.get(r.get("entry", -1)))
             for n, r in enumerate(st["rows"])])
    if source_bytes:
        FILES_DIR.mkdir(parents=True, exist_ok=True)
        ext = ".xlsm" if source_name.lower().endswith(".xlsm") else ".xlsx"
        path = FILES_DIR / f"v{vid}{ext}"
        path.write_bytes(source_bytes)
        conn.execute("UPDATE mep_library_versions SET source_file=? WHERE id=?", (str(path), vid))


def sheets(conn, vid: int) -> List[Dict[str, Any]]:
    out = []
    for r in conn.execute("SELECT * FROM mep_library_sheets WHERE version_id=? ORDER BY sheet_index", (vid,)):
        n = conn.execute("SELECT COUNT(*) FROM mep_library_rows lr JOIN mep_libraries l ON l.id=lr.item_id WHERE lr.sheet_id=?",
                         (r["id"],)).fetchone()[0]
        if not json.loads(r["mapping_json"] or "{}"):  # sheet giữ nguyên (không lấy giá): đếm số dòng
            n = conn.execute("SELECT COUNT(*) FROM mep_library_rows WHERE sheet_id=?", (r["id"],)).fetchone()[0]
        out.append({"id": r["id"], "name": r["sheet_name"], "header_row": r["header_row"], "first_col": r["first_col"],
                    "last_col": r["last_col"], "headers": json.loads(r["headers_json"] or "[]"),
                    "mapping": json.loads(r["mapping_json"] or "{}"), "count": n})
    return out


def sheet_rows(conn, sheet_id: int, items_by_id: Dict[int, Dict[str, Any]], filtered: bool) -> List[Dict[str, Any]]:
    """Dòng theo đúng thứ tự file. Dòng mục giá gắn với bản ghi thư viện hiện tại (đã sửa thì hiện giá trị mới).
    Khi đang lọc/tìm: chỉ trả các dòng mục khớp (items_by_id đã lọc)."""
    rows = []
    for r in conn.execute("SELECT * FROM mep_library_rows WHERE sheet_id=? ORDER BY seq, id", (sheet_id,)):
        it = items_by_id.get(r["item_id"]) if r["item_id"] else None
        if r["kind"] == "item":
            if not it:
                if filtered:
                    continue
                rows.append({"id": r["id"], "row": r["excel_row"], "kind": "deleted", "cells": json.loads(r["cells_json"] or "[]")})
                continue
            rows.append({"id": r["id"], "row": r["excel_row"], "kind": "item", "cells": json.loads(r["cells_json"] or "[]"), "item": it})
        elif not filtered:
            rows.append({"id": r["id"], "row": r["excel_row"], "kind": r["kind"], "cells": json.loads(r["cells_json"] or "[]")})
    return rows


def link_new_item(conn, vid: int, sheet_id: int, item_id: int):
    seq = conn.execute("SELECT COALESCE(MAX(seq), 0) + 1 FROM mep_library_rows WHERE sheet_id=?", (sheet_id,)).fetchone()[0]
    conn.execute("INSERT INTO mep_library_rows (version_id, sheet_id, excel_row, seq, kind, cells_json, item_id) VALUES (?,?,?,?,?,?,?)",
                 (vid, sheet_id, None, seq, "item", "[]", item_id))


def _num_equal(a: Any, b: Any) -> bool:
    try:
        return a is not None and b is not None and abs(float(a) - float(b)) < 1e-6
    except (TypeError, ValueError):
        return False


def export_original(vid: int) -> Optional[Dict[str, Any]]:
    """Trả file gốc đã nạp, với các ô đã sửa trên web được ghi đè; mục thêm mới nối cuối sheet;
    mục đã xóa thì xóa trắng dòng (không xóa dòng để không làm lệch công thức). None nếu không có file gốc."""
    with get_connection() as conn:
        v = conn.execute("SELECT * FROM mep_library_versions WHERE id=?", (vid,)).fetchone()
        if not v or not v["source_file"] or not Path(v["source_file"]).exists():
            return None
        sh = sheets(conn, vid)
        items = {r["id"]: dict(r) for r in conn.execute("SELECT * FROM mep_libraries WHERE version=?", (vid,))}
        all_rows = {s["id"]: [dict(r) for r in conn.execute("SELECT * FROM mep_library_rows WHERE sheet_id=? ORDER BY seq, id", (s["id"],))] for s in sh}
    src = Path(v["source_file"])
    wb = openpyxl.load_workbook(src, keep_vba=src.suffix.lower() == ".xlsm")
    changed_cells = 0
    for s in sh:
        if s["name"] not in wb.sheetnames:
            continue
        ws = wb[s["name"]]
        mapping, first = s["mapping"], s["first_col"]
        last_row = max([r["excel_row"] or 0 for r in all_rows[s["id"]]] + [s["header_row"]])
        for r in all_rows[s["id"]]:
            cells = json.loads(r["cells_json"] or "[]")
            it = items.get(r["item_id"]) if r["item_id"] else None
            if r["kind"] != "item":
                continue
            if r["excel_row"] is None:  # mục thêm trên web
                if not it:
                    continue
                last_row += 1
                for f, col in mapping.items():
                    if f in EDITABLE and it.get(f) not in (None, ""):
                        ws.cell(row=last_row, column=col, value=it[f])
                        changed_cells += 1
                continue
            if not it:  # mục đã xóa trên web
                for col in range(s["first_col"], s["last_col"] + 1):
                    ws.cell(row=r["excel_row"], column=col).value = None
                changed_cells += 1
                continue
            for f, col in mapping.items():
                if f not in EDITABLE:
                    continue
                orig = cells[col - first] if 0 <= col - first < len(cells) else None
                new = it.get(f)
                if f == "waste_percent" and isinstance(orig, (int, float)) and 0 < orig < 1 and new is not None:
                    new = new / 100  # ô gốc định dạng %
                same = _num_equal(orig, new) if isinstance(new, (int, float)) else (str(orig or "").strip() == str(new or "").strip())
                cur = ws.cell(row=r["excel_row"], column=col).value
                if (not same and f == "total_price" and isinstance(cur, str) and cur.startswith("=")
                        and any(k in mapping for k in ("material_price", "labor_price", "machine_price"))):
                    continue  # ô tổng là công thức cộng VT/NC/Máy: để Excel tự tính lại
                if not same:
                    ws.cell(row=r["excel_row"], column=col, value=new)
                    changed_cells += 1
    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return {"data": buf.getvalue(), "ext": src.suffix.lower(), "changed": changed_cells}

