"""Soát xét nhiều sheet trong một file NTP và giữ nguyên cấu trúc từng sheet (cột gốc, thứ tự dòng,
dòng nhóm / tổng) để hiển thị đúng như file người dùng nạp."""
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl

from backend import boq_reader, mep_router
from backend.library_service import cell_json
from backend.mep_db import get_connection
from backend.mep_excel import get_workbook_record
from backend.xlsx_utils import display_cell

PRICE_KEYS = ("price", "material_price", "labor_price", "machine_price")


def ensure_schema():
    with get_connection() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS mep_review_sheet_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            excel_row INTEGER NOT NULL,
            kind TEXT NOT NULL,
            cells_json TEXT,
            item_id INTEGER
        );
        CREATE INDEX IF NOT EXISTS idx_rev_sheet_rows ON mep_review_sheet_rows(session_id, sheet_name, excel_row);
        """)
        cols = {r[1] for r in conn.execute("PRAGMA table_info(mep_packages)")}
        if "sheets_json" not in cols:
            conn.execute("ALTER TABLE mep_packages ADD COLUMN sheets_json TEXT")
        if "custom_cols_json" not in cols:  # cột soát xét do người dùng tạo thêm
            conn.execute("ALTER TABLE mep_packages ADD COLUMN custom_cols_json TEXT")
        rcols = {r[1] for r in conn.execute("PRAGMA table_info(mep_review_sheet_rows)")}
        if "fcells_json" not in rcols:  # công thức gốc của từng ô (hiện trên thanh fx như Excel)
            conn.execute("ALTER TABLE mep_review_sheet_rows ADD COLUMN fcells_json TEXT")
        if "pos" not in rcols:  # vị trí hiển thị (dòng KSCP thêm/chèn nằm giữa các dòng file)
            conn.execute("ALTER TABLE mep_review_sheet_rows ADD COLUMN pos REAL")
        if "round_unit" not in cols:  # làm tròn đơn giá soát xét (100 = …00; 0 = không)
            conn.execute("ALTER TABLE mep_packages ADD COLUMN round_unit REAL DEFAULT 100")
        # Trạng thái gói gọn còn 3: Đang soát xét / Đã phản hồi BCHCT / Xong (đổi trạng thái cũ một lần)
        conn.execute("UPDATE mep_packages SET status='Đang soát xét' WHERE status IN ('Chờ duyệt','Cần giảm')")
        conn.execute("UPDATE mep_packages SET status='Xong' WHERE status='Đã duyệt'")
        pcols = {r[1] for r in conn.execute("PRAGMA table_info(mep_projects)")}
        if pcols and "kh_region" not in pcols:  # vị trí dự án: xa = Xa biển, gan = Gần biển (tra bảng KH vật tư)
            conn.execute("ALTER TABLE mep_projects ADD COLUMN kh_region TEXT DEFAULT 'xa'")
        if "he_so" not in cols:  # hệ số của gói (vd vận hành 2 tháp ×1,8) – nhân vào đơn giá nhân công trực
            conn.execute("ALTER TABLE mep_packages ADD COLUMN he_so REAL DEFAULT 1")
        if "pa" not in cols:  # phương án: PA1 Dự toán / PA2 BCHCT soát xét / PA3 P.KSCP soát xét
            conn.execute("ALTER TABLE mep_packages ADD COLUMN pa TEXT DEFAULT 'PA2'")
        if "p2_order_json" not in cols:  # thứ tự cột Phần 2 (kéo lên/xuống trong Cấu hình cột)
            conn.execute("ALTER TABLE mep_packages ADD COLUMN p2_order_json TEXT")


def _sheet_names(workbook_id: str) -> List[str]:
    rec = get_workbook_record(workbook_id)
    info = rec.get("sheets_info")
    sheets = json.loads(info) if isinstance(info, str) and info else []
    return [s["name"] if isinstance(s, dict) else s for s in sheets]


def _is_boq(mapping: Dict[str, Any]) -> bool:
    return bool(mapping.get("description") and any(mapping.get(k) for k in PRICE_KEYS))


def _config(workbook_id: str, sheet: str, header_row: Optional[int] = None, mapping: Optional[Dict[str, Any]] = None,
            saved_headers: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    sug = mep_router.suggest_boq_mapping({"workbook_id": workbook_id, "sheet_name": sheet,
                                          "header_row": header_row or 1, "force_header_row": bool(header_row)})
    # Cấu hình lưu từ bản trước chưa có cột %KH → nhận diện lại theo tiêu đề
    if mapping and "depreciation_pct" not in mapping:
        mapping = None
    # Cột đã gán không còn đúng tiêu đề (vd "Đơn giá" đang trỏ vào cột "Thời gian bắt đầu") → nhận diện lại
    if mapping:
        from backend import boq_reader
        kws = dict(boq_reader.FIELDS)
        hdr = {int(c): t for c, t in sug["headers"].items()}
        for f, v in mapping.items():
            if v and f in ("price", "amount", "depreciation_pct") and hdr.get(int(v)) and _is_boq(sug["mapping"]) \
                    and boq_reader._score(f, hdr[int(v)], kws[f]) == 0 and sug["mapping"].get(f) and int(sug["mapping"][f]) != int(v):
                mapping = None
                break
    # Cấu hình cũ chưa có cột thời gian → thêm cột nhận diện được (nếu cột đó chưa dùng cho trường khác)
    if mapping:
        used = {int(v) for v in mapping.values() if v}
        mapping = dict(mapping)
        for k in ("start_date", "end_date", "dep_done_pct"):
            v = sug["mapping"].get(k)
            if k not in mapping and v and int(v) not in used:
                mapping[k] = v
    m = {k: (int(v) if v else None) for k, v in (mapping or sug["mapping"]).items()}
    headers = {int(c): t for c, t in sug["headers"].items()}
    cols = [c for c in list(headers) + [v for v in m.values() if v]]
    return {"sheet": sheet, "header_row": int(sug["header_row"]), "mapping": m, "headers": headers,
            "first_col": min(cols) if cols else 1, "last_col": max(cols) if cols else 1}


def resolve_configs(package: Dict[str, Any], body: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Cấu hình các sheet cần soát xét.
    - body có sheet_name + mapping: cập nhật/ thêm sheet đó, giữ các sheet khác.
    - body có remove_sheet: bỏ sheet đó.
    - còn lại: dùng cấu hình đã lưu; chưa có thì tự nhận diện MỌI sheet có bảng giá."""
    wb_id = package["workbook_id"]
    names = _sheet_names(wb_id)
    saved = [c for c in json.loads(package.get("sheets_json") or "[]") if c.get("sheet") in names]
    if body.get("remove_sheet"):
        saved = [c for c in saved if c["sheet"] != body["remove_sheet"]]
        if not saved:
            raise ValueError("Cần giữ ít nhất một sheet để soát xét")
        return [_config(wb_id, c["sheet"], c["header_row"], c["mapping"], c.get("headers")) for c in saved]
    if body.get("sheet_name") and body.get("mapping"):
        new = _config(wb_id, body["sheet_name"], body.get("header_row"), body["mapping"])
        if not _is_boq(new["mapping"]):
            raise ValueError("Cần chọn cột Hạng mục và ít nhất một cột Đơn giá")
        out = [c for c in saved if c["sheet"] != new["sheet"]] + [new]
        return sorted(out, key=lambda c: names.index(c["sheet"]))
    if saved:
        return [_config(wb_id, c["sheet"], c["header_row"], c["mapping"], c.get("headers")) for c in saved]
    found = []
    for name in names:
        cfg = _config(wb_id, name)
        if _is_boq(cfg["mapping"]):
            found.append(cfg)
    if not found:
        raise ValueError("Không tìm thấy sheet nào có cột Hạng mục và Đơn giá. Mở tab Cấu hình cột để chọn.")
    return found


def run_all(package: Dict[str, Any], configs: List[Dict[str, Any]], threshold: float) -> Dict[str, Any]:
    """Chạy soát xét từng sheet rồi gộp thành MỘT phiên của gói thầu."""
    main_sid: Optional[str] = None
    summary: Dict[str, float] = {}
    per_sheet = []
    for cfg in configs:
        res = mep_router.run_price_review({
            "workbook_id": package["workbook_id"], "sheet_name": cfg["sheet"], "header_row": cfg["header_row"],
            "mapping": cfg["mapping"], "threshold_percent": threshold, "version_id": package.get("version_id"),
        })
        per_sheet.append({"sheet": cfg["sheet"], "total": res["summary"]["total"]})
        for k, v in res["summary"].items():
            summary[k] = summary.get(k, 0) + v
        with get_connection() as conn:
            if main_sid is None:
                main_sid = res["session_id"]
            else:
                conn.execute("UPDATE mep_review_items SET session_id=? WHERE session_id=?", (main_sid, res["session_id"]))
                conn.execute("DELETE FROM mep_review_sessions WHERE id=?", (res["session_id"],))
    with get_connection() as conn:
        conn.execute("""UPDATE mep_review_sessions SET sheet_name=?, total_items=?, passed_count=?, higher_count=?, lower_count=?,
                        not_found_count=?, spec_mismatch_count=?, manual_check_count=?, total_variance=?, mapping_config=? WHERE id=?""",
                     (", ".join(c["sheet"] for c in configs), summary.get("total", 0), summary.get("passed", 0), summary.get("higher", 0),
                      summary.get("lower", 0), summary.get("not_found", 0), summary.get("spec_mismatch", 0),
                      summary.get("manual_check", 0), summary.get("total_variance", 0), json.dumps(configs[0]["mapping"]), main_sid))
    snapshot(main_sid, package["workbook_id"], configs)
    _p2_starts(main_sid, package["workbook_id"], configs)
    return {"session_id": main_sid, "summary": {k: (int(v) if k != "total_variance" else v) for k, v in summary.items()},
            "per_sheet": per_sheet}


def _p2_starts(session_id: str, workbook_id: str, configs: List[Dict[str, Any]]):
    """Cột bắt đầu khối soát xét khi xuất Excel (sau vùng đang dùng 1 cột trống) – web dùng cùng chữ cột
    để công thức gõ trên bảng (=M7*1,1) xuất ra file đúng ô."""
    from backend.review_export import _last_used_col
    rec = get_workbook_record(workbook_id)
    wb = openpyxl.load_workbook(Path(rec["file_path"]))
    try:
        with get_connection() as conn:
            for cfg in configs:
                rows = [r[0] for r in conn.execute("SELECT excel_row FROM mep_review_items WHERE session_id=? AND sheet_name=?"
                                                   " AND COALESCE(is_added,0)=0", (session_id, cfg["sheet"]))]
                last = max(rows) if rows else int(cfg.get("header_row") or 1)
                cfg["p2_start"] = _last_used_col(wb[cfg["sheet"]], list(range(1, last + 1))) + 2
    finally:
        wb.close()


def snapshot(session_id: str, workbook_id: str, configs: List[Dict[str, Any]]):
    """Chụp lại mọi dòng (kể cả dòng nhóm/tổng) của vùng bảng trong từng sheet, đúng thứ tự file."""
    rec = get_workbook_record(workbook_id)
    wb = openpyxl.load_workbook(Path(rec["file_path"]), data_only=True)
    try:
        with get_connection() as conn:
            items = {(r["sheet_name"], r["excel_row"]): r["id"] for r in
                     conn.execute("SELECT id, sheet_name, excel_row FROM mep_review_items WHERE session_id=? AND COALESCE(is_added,0)=0", (session_id,))}
            conn.execute("DELETE FROM mep_review_sheet_rows WHERE session_id=?", (session_id,))
            batch = []
            for cfg in configs:
                ws = wb[cfg["sheet"]]
                for r in range(cfg["header_row"] + 1, (ws.max_row or 0) + 1):
                    cells = [cell_json(display_cell(ws.cell(row=r, column=c))) for c in range(cfg["first_col"], cfg["last_col"] + 1)]
                    iid = items.get((cfg["sheet"], r))
                    if iid is None and all(v is None for v in cells):
                        continue
                    batch.append((session_id, cfg["sheet"], r, "item" if iid else "group",
                                  json.dumps(cells, ensure_ascii=False), iid))
            conn.executemany("INSERT INTO mep_review_sheet_rows (session_id, sheet_name, excel_row, kind, cells_json, item_id)"
                             " VALUES (?,?,?,?,?,?)", batch)
    finally:
        wb.close()
    capture_formulas(session_id, workbook_id, configs)


def capture_formulas(session_id: str, workbook_id: str, configs: List[Dict[str, Any]]):
    """Đọc công thức gốc (=SUBTOTAL(9,I6:I30), =E7*G7+F7…) của các ô trong vùng bảng, lưu theo dòng."""
    rec = get_workbook_record(workbook_id)
    wb = openpyxl.load_workbook(Path(rec["file_path"]))
    try:
        cfg_by = {c["sheet"]: c for c in configs if c.get("first_col") and c.get("last_col")}
        with get_connection() as conn:
            rows = conn.execute("SELECT id, sheet_name, excel_row FROM mep_review_sheet_rows WHERE session_id=?", (session_id,)).fetchall()
            upd = []
            for r in rows:
                cfg = cfg_by.get(r["sheet_name"])
                if not cfg or r["sheet_name"] not in wb.sheetnames:
                    upd.append(("[]", r["id"]))
                    continue
                ws = wb[r["sheet_name"]]
                fs = []
                for c in range(int(cfg["first_col"]), int(cfg["last_col"]) + 1):
                    v = ws.cell(row=r["excel_row"], column=c).value
                    fs.append(v if isinstance(v, str) and v.startswith("=") else None)
                upd.append((json.dumps(fs if any(fs) else [], ensure_ascii=False), r["id"]))
            conn.executemany("UPDATE mep_review_sheet_rows SET fcells_json=? WHERE id=?", upd)
    finally:
        wb.close()


def kh_infos(conn, package: Dict[str, Any], region: str) -> Dict[int, Dict[str, Any]]:
    """%KH soát xét theo thời gian cho từng hạng mục: số tháng = ROUNDUP((Kết thúc − Bắt đầu)/30,0) → tra bảng KH VẬT TƯ."""
    from backend import depreciation as dp
    sid = package.get("session_id")
    if not sid:
        return {}
    cfgs = {c["sheet"]: c for c in json.loads(package.get("sheets_json") or "[]")}
    if not any((c.get("mapping") or {}).get("start_date") and (c.get("mapping") or {}).get("end_date") for c in cfgs.values()):
        return {}
    tables = dp.load_tables(conn)
    # % đã KH + ngày sửa trên bảng (rv_extra d_start / d_end) của từng hạng mục
    extra: Dict[int, Dict[str, Any]] = {}
    for r in conn.execute("SELECT id, boq_dep_done, rv_extra FROM mep_review_items WHERE session_id=?", (sid,)):
        try:
            ex = json.loads(r["rv_extra"] or "{}")
        except ValueError:
            ex = {}
        extra[r["id"]] = {"done": r["boq_dep_done"], "d_start": ex.get("d_start"), "d_end": ex.get("d_end")}
    out: Dict[int, Dict[str, Any]] = {}
    for sheet, rows in structure(conn, sid).items():
        c = cfgs.get(sheet)
        if not c:
            continue
        m, fc = c.get("mapping") or {}, int(c.get("first_col") or 1)
        if not (m.get("start_date") and m.get("end_date")):
            continue
        cell = lambda r, f: (r["cells"][int(m[f]) - fc] if m.get(f) and 0 <= int(m[f]) - fc < len(r["cells"]) else None)
        parent = ""
        for r in rows:
            if not r.get("item_id"):
                parent = str(cell(r, "description") or "").strip() or parent
                continue
            ex = extra.get(r["item_id"], {})
            months = dp.months_between(ex.get("d_start") or cell(r, "start_date"), ex.get("d_end") or cell(r, "end_date"))
            info = dp.expected_pct(str(cell(r, "description") or ""), parent, months, tables, region)
            info["table"] = tables.get("version")
            # %KH soát xét = %KH theo bảng − % ĐÃ KH (phần đã khấu hao trước đó) → công thức =0,53-G7
            done = ex.get("done")
            if done is None and m.get("dep_done_pct"):
                from backend.mep_router import dep_fraction
                done = dep_fraction(cell(r, "dep_done_pct"))
            if m.get("dep_done_pct") and info.get("pct") is not None and r.get("row"):
                from openpyxl.utils import get_column_letter
                base = float(info["pct"])
                info["base"] = base
                info["done"] = float(done or 0)
                info["pct"] = max(0.0, round(base - float(done or 0), 6))
                info["formula"] = f"={round(base, 4):g}".replace(".", ",") + f"-{get_column_letter(int(m['dep_done_pct']))}{r['row']}"
            out[r["item_id"]] = info
    return out


def _set_dep_formula(conn, iid: int, formula: Optional[str]):
    """Ghi / bỏ công thức của ô %KH soát xét (rv_extra f_rv_dep_pct)."""
    r = conn.execute("SELECT rv_extra FROM mep_review_items WHERE id=?", (iid,)).fetchone()
    try:
        ex = json.loads((r["rv_extra"] if r else None) or "{}")
    except ValueError:
        ex = {}
    if formula:
        ex["f_rv_dep_pct"] = formula
    else:
        ex.pop("f_rv_dep_pct", None)
    conn.execute("UPDATE mep_review_items SET rv_extra=? WHERE id=?", (json.dumps(ex, ensure_ascii=False) if ex else None, iid))


def apply_kh(package: Dict[str, Any], region: str, only_if: Optional[Dict[int, Optional[float]]] = None) -> int:
    """Ghi %KH soát xét theo thời gian vào các dòng chưa sửa tay (hoặc đang bằng %KH tự tính trước đó: only_if)."""
    with get_connection() as conn:
        infos = kh_infos(conn, package, region)
        n = 0
        for iid, info in infos.items():
            if info.get("pct") is None:
                continue
            r = conn.execute("SELECT rv_dep_pct, rv_touched, is_added FROM mep_review_items WHERE id=?", (iid,)).fetchone()
            if not r or r["is_added"]:
                continue
            if only_if is not None:
                old = only_if.get(iid)
                if old is None or r["rv_dep_pct"] is None or abs(float(r["rv_dep_pct"]) - old) > 1e-9:
                    continue
            elif r["rv_touched"]:
                continue
            conn.execute("UPDATE mep_review_items SET rv_dep_pct=? WHERE id=?", (info["pct"], iid))
            _set_dep_formula(conn, iid, info.get("formula"))
            n += 1
    return n


def structure(conn, session_id: Optional[str]) -> Dict[str, List[Dict[str, Any]]]:
    out: Dict[str, List[Dict[str, Any]]] = {}
    if not session_id:
        return out
    for r in conn.execute("SELECT * FROM mep_review_sheet_rows WHERE session_id=? ORDER BY COALESCE(pos, excel_row), id", (session_id,)):
        out.setdefault(r["sheet_name"], []).append(
            {"row": r["excel_row"], "kind": r["kind"], "cells": json.loads(r["cells_json"] or "[]"), "item_id": r["item_id"],
             "sid": r["id"], "pos": r["pos"] if r["pos"] is not None else r["excel_row"],
             "f": json.loads(r["fcells_json"]) if r["fcells_json"] else None})
    return out


ADDED_FIELDS = ("excel_col", "boq_code", "boq_description", "boq_unit", "boq_quantity", "boq_price", "status", "boq_stt",
                "rv_quantity", "rv_material", "rv_labor", "rv_machine", "rv_total", "rv_note", "rv_dep_pct", "rv_extra",
                "rv_touched", "is_added", "updated_at")


def add_row(session_id: str, anchor_sid: int, where: str = "after") -> Dict[str, int]:
    """Thêm (after) / chèn (before) một dòng KSCP bổ sung cạnh dòng anchor của bảng. Trả id hạng mục mới."""
    with get_connection() as conn:
        a = conn.execute("SELECT * FROM mep_review_sheet_rows WHERE id=? AND session_id=?", (anchor_sid, session_id)).fetchone()
        if not a:
            raise ValueError("Không tìm thấy dòng để chèn cạnh")
        p = a["pos"] if a["pos"] is not None else a["excel_row"]
        if where == "before":
            nb = conn.execute("SELECT MAX(COALESCE(pos, excel_row)) FROM mep_review_sheet_rows WHERE session_id=? AND sheet_name=?"
                              " AND COALESCE(pos, excel_row) < ?", (session_id, a["sheet_name"], p)).fetchone()[0]
            new_pos = (p + nb) / 2 if nb is not None else p - 0.5
        else:
            nb = conn.execute("SELECT MIN(COALESCE(pos, excel_row)) FROM mep_review_sheet_rows WHERE session_id=? AND sheet_name=?"
                              " AND COALESCE(pos, excel_row) > ?", (session_id, a["sheet_name"], p)).fetchone()[0]
            new_pos = (p + nb) / 2 if nb is not None else p + 1
        cur = conn.execute("""INSERT INTO mep_review_items (session_id, sheet_name, excel_row, excel_col, boq_description, boq_quantity,
                              boq_price, status, rv_quantity, rv_touched, is_added, updated_at)
                              VALUES (?,?,?,?,?,?,?,?,?,?,?,datetime('now'))""",
                           (session_id, a["sheet_name"], a["excel_row"], 0, "", 0, 0, "not_found", 0, 1, 1))
        iid = cur.lastrowid
        sid = conn.execute("INSERT INTO mep_review_sheet_rows (session_id, sheet_name, excel_row, kind, cells_json, item_id, pos)"
                           " VALUES (?,?,?,?,?,?,?)", (session_id, a["sheet_name"], a["excel_row"], "item", "[]", iid, new_pos)).lastrowid
    return {"id": iid, "sid": sid}


def delete_added_row(session_id: str, item_id: int):
    with get_connection() as conn:
        r = conn.execute("SELECT is_added FROM mep_review_items WHERE id=? AND session_id=?", (item_id, session_id)).fetchone()
        if not r:
            raise ValueError("Không tìm thấy dòng")
        if not r["is_added"]:
            raise ValueError("Chỉ xóa được dòng KSCP thêm vào; dòng của file NTP giữ nguyên")
        conn.execute("DELETE FROM mep_review_sheet_rows WHERE session_id=? AND item_id=?", (session_id, item_id))
        conn.execute("DELETE FROM mep_review_items WHERE id=?", (item_id,))


def carry_added_rows(old_sid: Optional[str], new_sid: str) -> int:
    """Soát xét lại → chép các dòng KSCP thêm vào sang phiên mới (giữ vị trí, giá, lịch sử)."""
    if not old_sid or old_sid == new_sid:
        return 0
    n = 0
    with get_connection() as conn:
        rows = conn.execute("""SELECT i.*, s.pos AS s_pos FROM mep_review_items i JOIN mep_review_sheet_rows s
                               ON s.item_id = i.id AND s.session_id = i.session_id
                               WHERE i.session_id=? AND COALESCE(i.is_added,0)=1""", (old_sid,)).fetchall()
        for r in rows:
            cols = ", ".join(ADDED_FIELDS)
            cur = conn.execute(f"INSERT INTO mep_review_items (session_id, sheet_name, excel_row, {cols}) VALUES (?,?,?,{','.join('?' * len(ADDED_FIELDS))})",
                               (new_sid, r["sheet_name"], r["excel_row"], *[r[k] for k in ADDED_FIELDS]))
            conn.execute("INSERT INTO mep_review_sheet_rows (session_id, sheet_name, excel_row, kind, cells_json, item_id, pos)"
                         " VALUES (?,?,?,?,?,?,?)", (new_sid, r["sheet_name"], r["excel_row"], "item", "[]", cur.lastrowid, r["s_pos"]))
            conn.execute("UPDATE mep_review_edits SET item_id=? WHERE item_id=?", (cur.lastrowid, r["id"]))
            n += 1
    return n


def public_configs(package: Dict[str, Any]) -> List[Dict[str, Any]]:
    cfgs = json.loads(package.get("sheets_json") or "[]")
    if not cfgs and package.get("sheet_name") and package.get("mapping_json"):
        cfgs = [{"sheet": package["sheet_name"], "header_row": package.get("header_row") or 1,
                 "mapping": json.loads(package["mapping_json"]), "headers": {}, "first_col": None, "last_col": None}]
    for c in cfgs:
        c["headers"] = {str(k): v for k, v in (c.get("headers") or {}).items()}
    return cfgs
