"""Dự án, gói thầu và version thư viện giá cho giao diện Soát xét giá."""
import io
import json
import re
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from backend.mep_db import get_connection, now_iso
from backend.mep_engine import normalize_mep_description
from backend.mep_excel import UPLOADS_DIR, get_workbook_record, save_uploaded_workbook
from backend import mep_router
from backend import library_service, review_service, review_export, review_sheets, review_versions
from backend.xlsx_utils import bytes_response
from fastapi import Request
from urllib.parse import unquote
from backend.xlsx_utils import write_table, xlsx_response

router = APIRouter(prefix="/api/mep", tags=["MEP Packages"])

STATUSES = ["Nháp", "Đang soát xét", "Đã phản hồi BCHCT", "Xong"]  # trạng thái gói thầu (Nháp: chưa soát xét)


def init_package_tables():
    with get_connection() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS mep_projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            location TEXT,
            created_at TEXT
        );
        CREATE TABLE IF NOT EXISTS mep_packages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            work_type TEXT DEFAULT 'Điện',
            subcontractor TEXT,
            version_id INTEGER,
            deadline TEXT,
            status TEXT DEFAULT 'Nháp',
            workbook_id TEXT,
            session_id TEXT,
            sheet_name TEXT,
            vat_percent REAL DEFAULT 10,
            note TEXT,
            created_at TEXT,
            updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS mep_library_versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            effective_date TEXT,
            scope TEXT DEFAULT 'Toàn bộ (Điện + Nước)',
            note TEXT,
            status TEXT DEFAULT 'draft',
            item_count INTEGER DEFAULT 0,
            error_count INTEGER DEFAULT 0,
            created_by TEXT,
            created_at TEXT
        );
        """)
        cols = {r[1] for r in conn.execute("PRAGMA table_info(mep_libraries)")}
        for name, typ in (("category", "TEXT"), ("waste_percent", "REAL")):
            if name not in cols:
                conn.execute(f"ALTER TABLE mep_libraries ADD COLUMN {name} {typ}")
        # Thư viện cũ (version=1) được gom vào một version hiện hành để không mất dữ liệu.
        has_versions = conn.execute("SELECT COUNT(*) FROM mep_library_versions").fetchone()[0]
        legacy = conn.execute("SELECT COUNT(*) FROM mep_libraries").fetchone()[0]
        if not has_versions and legacy:
            conn.execute(
                "INSERT INTO mep_library_versions (id, name, effective_date, scope, note, status, item_count, created_by, created_at)"
                " VALUES (1, ?, ?, 'Toàn bộ (Điện + Nước)', 'Thư viện có sẵn trước khi chia version', 'active', ?, 'Hệ thống', ?)",
                ("v" + now_iso()[:7].replace("-", "."), now_iso()[:10], legacy, now_iso()),
            )
            conn.execute("UPDATE mep_libraries SET version = 1 WHERE version IS NULL OR version NOT IN (SELECT id FROM mep_library_versions)")


init_package_tables()
library_service.ensure_library_schema()
review_service.ensure_schema()
review_sheets.ensure_schema()
review_versions.ensure_schema()


# ------------------------------------------------------------------ helpers

def _num(v) -> Optional[float]:
    return mep_router._num(v)


def _active_version_id(conn) -> Optional[int]:
    r = conn.execute("SELECT id FROM mep_library_versions WHERE status='active' ORDER BY id DESC LIMIT 1").fetchone()
    return r[0] if r else None


def _session_items(conn, session_id: Optional[str]) -> List[Dict[str, Any]]:
    return review_service.session_items(conn, session_id)


def _package_dict(conn, row) -> Dict[str, Any]:
    p = dict(row)
    items = _session_items(conn, p.get("session_id"))
    t = review_service.totals_k3(items) if p.get("pa") == "PA3" else review_service.totals(items)
    ntp, rev = t["ntp"], t["review"]
    p["item_count"] = len(items)
    p["counts"] = t["counts"]
    p["ntp_total"] = ntp if items else None
    p["review_total"] = rev if items and p.get("status") != "Nháp" else None
    p["diff_percent"] = round((rev - ntp) / ntp * 100, 1) if items and ntp and p["review_total"] is not None else None
    proj = conn.execute("SELECT name, location FROM mep_projects WHERE id=?", (p["project_id"],)).fetchone()
    p["project_name"] = proj["name"] if proj else ""
    p["project_location"] = proj["location"] if proj else ""
    ver = conn.execute("SELECT name, status FROM mep_library_versions WHERE id=?", (p.get("version_id"),)).fetchone() if p.get("version_id") else None
    p["version_name"] = ver["name"] if ver else ""
    if p.get("workbook_id"):
        wb = conn.execute("SELECT original_filename FROM mep_workbooks WHERE id=?", (p["workbook_id"],)).fetchone()
        p["file_name"] = wb["original_filename"] if wb else ""
    return p


_xlsx_response = xlsx_response
_write_table = write_table


# ------------------------------------------------------------------ projects

@router.get("/projects")
def list_projects():
    with get_connection() as conn:
        rows = conn.execute("SELECT * FROM mep_projects ORDER BY id").fetchall()
        out = []
        for r in rows:
            d = dict(r)
            pk = conn.execute("SELECT id, code, name, status FROM mep_packages WHERE project_id=? ORDER BY code", (r["id"],)).fetchall()
            d["packages"] = [dict(x) for x in pk]
            out.append(d)
        return out


@router.post("/projects")
def create_project(body: Dict[str, Any]):
    name = str(body.get("name") or "").strip()
    if not name:
        raise HTTPException(400, "Nhập tên dự án")
    with get_connection() as conn:
        if conn.execute("SELECT 1 FROM mep_projects WHERE name=?", (name,)).fetchone():
            raise HTTPException(400, "Dự án đã tồn tại")
        region = "gan" if body.get("kh_region") == "gan" else "xa"
        cur = conn.execute("INSERT INTO mep_projects (name, location, created_at, kh_region) VALUES (?,?,?,?)",
                           (name, str(body.get("location") or "").strip(), now_iso(), region))
        return {"id": cur.lastrowid, "name": name}


def _project_region(conn, project_id) -> str:
    r = conn.execute("SELECT kh_region FROM mep_projects WHERE id=?", (project_id,)).fetchone()
    return (r["kh_region"] if r and r["kh_region"] in ("xa", "gan") else "xa")


@router.put("/projects/{project_id}/region")
def set_project_region(project_id: int, body: Dict[str, Any]):
    """Đổi vị trí dự án (Xa biển / Gần biển) → %KH soát xét các gói tính lại theo bảng tương ứng
    (chỉ ô %KH đang là số tự tính, ô đã sửa tay giữ nguyên)."""
    region = "gan" if body.get("kh_region") == "gan" else "xa"
    with get_connection() as conn:
        old = _project_region(conn, project_id)
        conn.execute("UPDATE mep_projects SET kh_region=? WHERE id=?", (region, project_id))
        pkgs = [dict(r) for r in conn.execute("SELECT * FROM mep_packages WHERE project_id=?", (project_id,))]
    n = 0
    for p in pkgs:
        with get_connection() as conn:
            prev = {i: x.get("pct") for i, x in review_sheets.kh_infos(conn, p, old).items()}
        n += review_sheets.apply_kh(p, region, only_if=prev)
    return {"kh_region": region, "updated": n}


@router.delete("/projects/{project_id}")
def delete_project(project_id: int, force: bool = False):
    """Xóa dự án. Còn gói thầu thì chỉ xóa khi force=true (xóa luôn các gói thầu của dự án;
    file báo giá đã nạp và lịch sử soát xét vẫn còn trong uploads/SQLite để truy vết)."""
    with get_connection() as conn:
        if not conn.execute("SELECT 1 FROM mep_projects WHERE id=?", (project_id,)).fetchone():
            raise HTTPException(404, "Không tìm thấy dự án")
        n = conn.execute("SELECT COUNT(*) FROM mep_packages WHERE project_id=?", (project_id,)).fetchone()[0]
        if n and not force:
            raise HTTPException(400, f"Dự án còn {n} gói thầu")
        conn.execute("DELETE FROM mep_packages WHERE project_id=?", (project_id,))
        conn.execute("DELETE FROM mep_projects WHERE id=?", (project_id,))
    return {"deleted": True, "packages_deleted": n}


# ------------------------------------------------------------------ packages

@router.get("/packages")
def list_packages(q: str = "", project_id: Optional[int] = None, work_type: str = "", subcontractor: str = "",
                  status: str = "", version_id: Optional[int] = None, date_from: str = "", date_to: str = ""):
    with get_connection() as conn:
        sql, params = "SELECT * FROM mep_packages WHERE 1=1", []
        if project_id:
            sql += " AND project_id=?"; params.append(project_id)
        if work_type:
            types = [t for t in work_type.split(",") if t]
            sql += f" AND work_type IN ({','.join('?' * len(types))})"; params += types
        if subcontractor:
            sql += " AND subcontractor=?"; params.append(subcontractor)
        if status:
            sql += " AND status=?"; params.append(status)
        if version_id:
            sql += " AND version_id=?"; params.append(version_id)
        if date_from:
            sql += " AND substr(created_at,1,10) >= ?"; params.append(date_from)
        if date_to:
            sql += " AND substr(created_at,1,10) <= ?"; params.append(date_to)
        if q.strip():
            sql += " AND (code LIKE ? OR name LIKE ? OR subcontractor LIKE ?)"
            params += [f"%{q.strip()}%"] * 3
        rows = conn.execute(sql + " ORDER BY code", params).fetchall()
        items = [_package_dict(conn, r) for r in rows]
        subs = [r[0] for r in conn.execute("SELECT DISTINCT subcontractor FROM mep_packages WHERE subcontractor<>'' ORDER BY 1")]
        return {"items": items, "subcontractors": subs, "statuses": STATUSES}


@router.get("/packages/export")
def export_packages():
    data = list_packages()["items"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Danh sách gói thầu"
    ws["A1"] = "DANH SÁCH GÓI THẦU SOÁT XÉT GIÁ"; ws["A1"].font = Font(bold=True, size=14, color="1E3A8A")
    rows = [[p["code"], p["name"], p["project_name"], p["subcontractor"], p["work_type"],
             p["ntp_total"], p["review_total"], p["diff_percent"], p["version_name"], p["status"], p["deadline"]] for p in data]
    _write_table(ws, 3, ["Mã gói", "Tên gói thầu", "Dự án", "Nhà thầu phụ", "Loại", "Giá NTP báo (đ)",
                         "Giá soát xét (đ)", "Chênh lệch (%)", "Version giá", "Trạng thái", "Hạn soát xét"],
                 rows, [14, 36, 24, 26, 10, 18, 18, 12, 12, 14, 14])
    return _xlsx_response(wb, "Danh-sach-goi-thau.xlsx")


# Phương án của gói
PA_LABELS = {"PA1": "Dự toán", "PA2": "BCHCT soát xét", "PA3": "P.KSCP soát xét"}
# PA1 – Dự toán: file bắt buộc đủ 13 cột A→M (dữ liệu nhập ở A–D, I, J; tool áp giá E, F và %KH H)
PA1_COLS = [("stt", "A – STT"), ("description", "B – Mô tả công việc"), ("unit", "C – Đơn vị"), ("quantity", "D – Khối lượng"),
            ("material_price", "E – Đơn giá VT"), ("labor_price", "F – Đơn giá NC"), ("dep_done_pct", "G – % Đã KH"),
            ("depreciation_pct", "H – % KH"), ("start_date", "I – Bắt đầu"), ("end_date", "J – Kết thúc"),
            ("price", "K – Đơn giá tổng"), ("amount", "L – Thành tiền"), ("note", "M – Ghi chú")]


def _check_pa1(configs: List[Dict[str, Any]]):
    for c in configs:
        m = c.get("mapping") or {}
        miss = [lbl for f, lbl in PA1_COLS if not m.get(f)]
        if miss:
            raise HTTPException(400, f"PA1 – Dự toán: sheet {c.get('sheet')} thiếu cột: {', '.join(miss)}. "
                                     "File dự toán phải có đủ các cột A → M")


def _fill_estimate(pid: int, p: Dict[str, Any], region: str) -> int:
    """PA1 – Dự toán: áp giá thư viện (Mặc định) vào Đơn giá VT / NC (cột E, F) và %KH theo thời gian (cột H, từ I và J)
    cho cả Phần 1 và Phần 2. Dòng không có giá thư viện → 0."""
    from backend import price_sources
    with get_connection() as conn:
        items = [i for i in review_service.session_items(conn, p["session_id"]) if not i.get("is_added")]
        s = conn.execute("SELECT threshold_percent FROM mep_review_sessions WHERE id=?", (p["session_id"],)).fetchone()
        pr = price_sources.prices(conn, items, "default", float((s["threshold_percent"] if s else None) or 10))["prices"]
        kh = review_sheets.kh_infos(conn, p, region)
    rows = []
    for it in items:
        if not (it.get("boq_description") or "").strip():
            continue
        x = pr.get(str(it["id"]))
        if x and any(x.get(k) is not None for k in ("material", "labor")):
            mat, lab = x.get("material") or 0, x.get("labor") or 0
        else:
            mat, lab = (x["total"] if x else 0), 0
        r = {"id": it["id"], "boq_material": mat, "boq_labor": lab, "rv_material": mat, "rv_labor": lab, "rv_total": None}
        k = kh.get(it["id"]) or {}
        if k.get("pct") is not None:
            r.update(boq_dep_pct=k["pct"], rv_dep_pct=k["pct"])
            if k.get("formula"):
                r.update(f_boq_dep_pct=k["formula"], f_rv_dep_pct=k["formula"])
        rows.append(r)
    if rows:
        review_service.save_rows(pid, p["session_id"], rows, "Tự động (PA1 – Dự toán)")
    return len(rows)


def _auto_review(conn_unused, package: Dict[str, Any], sheet_name: Optional[str] = None,
                 header_row: Optional[int] = None, mapping: Optional[Dict[str, Any]] = None,
                 threshold: float = 10.0) -> Dict[str, Any]:
    """Nhận diện cột (nếu chưa có) rồi chạy soát xét theo version thư viện của gói."""
    rec = get_workbook_record(package["workbook_id"])
    sheets = json.loads(rec.get("sheets_info") or "[]") if isinstance(rec.get("sheets_info"), str) else rec.get("sheets", [])
    names = [s["name"] if isinstance(s, dict) else s for s in sheets]
    sheet = sheet_name or package.get("sheet_name")
    if not mapping:
        # Chọn sheet có tiêu đề cột giống BOQ nhất (có Diễn giải + Đơn giá).
        best = None
        for name in ([sheet] if sheet else names):
            sug = mep_router.suggest_boq_mapping({"workbook_id": package["workbook_id"], "sheet_name": name, "header_row": header_row or 1})
            m = sug["mapping"]
            score = sum(1 for v in m.values() if v) + (5 if m.get("description") and (m.get("price") or m.get("material_price")) else 0)
            if not best or score > best[0]:
                best = (score, name, sug)
        if not best:
            raise HTTPException(400, "File không có sheet nào")
        _, sheet, sug = best
        mapping, header_row = sug["mapping"], sug["header_row"]
    res = mep_router.run_price_review({
        "workbook_id": package["workbook_id"], "sheet_name": sheet, "header_row": header_row,
        "mapping": mapping, "threshold_percent": threshold, "version_id": package.get("version_id"),
    })
    return {"session_id": res["session_id"], "sheet_name": sheet, "summary": res["summary"],
            "header_row": header_row, "mapping": mapping}


def _auto_code(conn, project_id: int) -> str:
    """Mã gói tự sinh: chữ đầu tên dự án + số thứ tự (VD: RT-01)."""
    proj = conn.execute("SELECT name FROM mep_projects WHERE id=?", (project_id,)).fetchone()
    words = [w for w in library_service.strip_accents(proj["name"] if proj else "").upper().split()
             if w.isalnum() and w not in ("CHUNG", "CU", "DU", "AN")]
    prefix = "".join(w[0] for w in words)[:4] or "GT"
    n = 1
    while conn.execute("SELECT 1 FROM mep_packages WHERE code=?", (f"{prefix}-{n:02d}",)).fetchone():
        n += 1
    return f"{prefix}-{n:02d}"


async def _reference_version(ref_file: UploadFile, code: str, user: str) -> int:
    """File giá đối chiếu riêng của gói: lưu thành version loại 'reference' (không hiện ở Thư viện giá)."""
    name = (ref_file.filename or "").lower()
    if not name.endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "File đối chiếu phải là .xlsx hoặc .xlsm")
    data = await ref_file.read()
    tmp = UPLOADS_DIR / f"ref_{code}_{now_iso()[:19].replace(':', '')}.xlsx"
    tmp.write_bytes(data)
    try:
        parsed = library_service.parse_library_file(tmp)
        v = library_service.create_version(f"Đối chiếu {code} · {ref_file.filename}", parsed["entries"],
                                           note=f"File đối chiếu riêng của gói {code}", created_by=user,
                                           status="reference", error_count=len(parsed["errors"]),
                                           structure=parsed["structure"], source_bytes=data, source_name=ref_file.filename or "")
    except ValueError as e:
        raise HTTPException(400, f"File đối chiếu: {e}")
    finally:
        tmp.unlink(missing_ok=True)
    return v["id"]


@router.post("/packages")
async def create_package(request: Request, project_id: int = Form(...), name: str = Form(...), code: str = Form(""),
                         work_type: str = Form("Điện"), subcontractor: str = Form(""),
                         version_id: Optional[int] = Form(None), deadline: str = Form(""),
                         start: bool = Form(False), pa: str = Form("PA2"), kh_region: str = Form(""), file: Optional[UploadFile] = File(None),
                         ref_file: Optional[UploadFile] = File(None)):
    name = name.strip()
    if not name:
        raise HTTPException(400, "Nhập tên gói thầu")
    pa = pa if pa in PA_LABELS else "PA2"
    if kh_region in ("xa", "gan"):  # vị trí dự án chọn ngay khi tạo gói
        set_project_region(project_id, {"kh_region": kh_region})
    with get_connection() as conn:
        if not conn.execute("SELECT 1 FROM mep_projects WHERE id=?", (project_id,)).fetchone():
            raise HTTPException(400, "Dự án không tồn tại")
        code = code.strip() or _auto_code(conn, project_id)
        if conn.execute("SELECT 1 FROM mep_packages WHERE code=?", (code,)).fetchone():
            raise HTTPException(400, f"Mã gói {code} đã tồn tại")
        if not version_id:
            version_id = _active_version_id(conn)
    if ref_file is not None and ref_file.filename:
        version_id = await _reference_version(ref_file, code, unquote(request.headers.get("X-CCM-User", ""))[:80])
    workbook_id = None
    if file is not None and file.filename:
        content = await file.read()
        try:
            workbook_id = save_uploaded_workbook(content, file.filename)["id"]
        except ValueError as e:
            raise HTTPException(400, str(e))
    with get_connection() as conn:
        cur = conn.execute(
            """INSERT INTO mep_packages (project_id, code, name, work_type, subcontractor, version_id, deadline,
                   status, workbook_id, created_at, updated_at, pa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (project_id, code, name, work_type, subcontractor.strip(), version_id, deadline, "Nháp", workbook_id, now_iso(), now_iso(), pa))
        pid = cur.lastrowid
    warning = None
    if start and workbook_id:
        try:
            run_package_review(pid, {})
        except HTTPException as e:
            warning = f"Đã tạo gói nhưng chưa soát xét tự động được: {e.detail}. Mở gói để chọn cột."
    return {"id": pid, "code": code, "warning": warning}


def _upgrade_stale_review(pid: int):
    """Gói soát xét từ bản cũ (cấu hình cột chưa có % KH vật tư) → tự soát xét lại một lần,
    giữ các dòng đã sửa tay, để cột % KH được nhận diện và tính đúng mà không cần bấm Tải dữ liệu."""
    with get_connection() as conn:
        row = conn.execute("SELECT session_id, workbook_id, sheets_json FROM mep_packages WHERE id=?", (pid,)).fetchone()
        if not row or not row["session_id"] or not row["workbook_id"] or not row["sheets_json"]:
            return
        cfgs = json.loads(row["sheets_json"] or "[]")
        def _fresh(c):  # cấu hình khớp tiêu đề file: cột Đơn giá / Thành tiền / thời gian đúng chỗ
            from backend import boq_reader
            hdr = {int(k): v for k, v in (c.get("headers") or {}).items()}
            if not hdr:
                return True
            sug = boq_reader.map_columns(hdr)
            m = c.get("mapping") or {}
            kws = dict(boq_reader.FIELDS)
            if any(sug.get(f) and f not in m for f in ("start_date", "end_date", "dep_done_pct")):
                return False
            return not any(m.get(f) and hdr.get(int(m[f])) and boq_reader._score(f, hdr[int(m[f])], kws[f]) == 0
                           and sug.get(f) and sug[f] != m[f] for f in ("price", "amount", "depreciation_pct"))
        if not cfgs or all("depreciation_pct" in (c.get("mapping") or {}) and c.get("p2_start") and _fresh(c) for c in cfgs):
            return
        s = conn.execute("SELECT threshold_percent FROM mep_review_sessions WHERE id=?", (row["session_id"],)).fetchone()
    try:
        run_package_review(pid, {"threshold_percent": (s["threshold_percent"] if s else None) or 10})
    except Exception:  # file gốc mất / lỗi đọc: vẫn mở gói như cũ
        pass


@router.get("/packages/{pid}")
def get_package(pid: int):
    _upgrade_stale_review(pid)
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM mep_packages WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "Không tìm thấy gói thầu")
        p = _package_dict(conn, row)
        p["items"] = _session_items(conn, p.get("session_id"))
        p["sheet_configs"] = review_sheets.public_configs(p)
        p["custom_columns"] = review_service.custom_columns(conn, pid)
        p["p2_order"] = review_service.p2_order(conn, pid, p["custom_columns"])
        p["round_unit"] = p.get("round_unit") if p.get("round_unit") is not None else 100
        if p.get("session_id") and p.get("workbook_id") and conn.execute(
                "SELECT 1 FROM mep_review_sheet_rows WHERE session_id=? AND fcells_json IS NULL LIMIT 1", (p["session_id"],)).fetchone():
            try:  # gói soát xét từ bản cũ: đọc bổ sung công thức gốc một lần
                review_sheets.capture_formulas(p["session_id"], p["workbook_id"], json.loads(p.get("sheets_json") or "[]"))
            except Exception:
                pass
        p["structure"] = review_sheets.structure(conn, p.get("session_id"))
        p["kh_region"] = _project_region(conn, p.get("project_id"))
        try:
            p["kh_info"] = {str(k): v for k, v in review_sheets.kh_infos(conn, p, p["kh_region"]).items()}
        except Exception:
            p["kh_info"] = {}
        if p.get("session_id"):
            s = conn.execute("SELECT * FROM mep_review_sessions WHERE id=?", (p["session_id"],)).fetchone()
            p["session"] = dict(s) if s else None
        if p.get("workbook_id"):
            rec = get_workbook_record(p["workbook_id"])
            info = rec.get("sheets_info")
            sheets = json.loads(info) if isinstance(info, str) and info else rec.get("sheets", [])
            p["sheets"] = [s["name"] if isinstance(s, dict) else s for s in sheets]
        return p


def _session_of(pid: int) -> str:
    with get_connection() as conn:
        row = conn.execute("SELECT session_id FROM mep_packages WHERE id=?", (pid,)).fetchone()
    if not row:
        raise HTTPException(404, "Không tìm thấy gói thầu")
    if not row["session_id"]:
        raise HTTPException(400, "Gói chưa soát xét")
    return row["session_id"]


@router.post("/packages/{pid}/rows")
def add_review_row(pid: int, body: Dict[str, Any], request: Request):
    """Thêm dòng (dưới) / chèn dòng (trên) dòng đang chọn – dòng KSCP bổ sung, không có giá NTP."""
    sid = _session_of(pid)
    try:
        new = review_sheets.add_row(sid, int(body.get("anchor_sid") or 0), "before" if body.get("where") == "before" else "after")
    except (ValueError, TypeError) as e:
        raise HTTPException(400, str(e))
    # Lịch sử chỉ ghi khi bấm Lưu: thao tác thêm dòng được gửi kèm lần Lưu kế tiếp (events)
    return new


@router.delete("/packages/{pid}/rows/{iid}")
def delete_review_row(pid: int, iid: int, request: Request):
    sid = _session_of(pid)
    try:
        review_sheets.delete_added_row(sid, iid)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {"deleted": True}


@router.put("/packages/{pid}/columns")
def save_custom_columns(pid: int, body: Dict[str, Any]):
    """Lưu danh sách cột soát xét tự tạo (tab Cấu hình cột → phần P.KSCP soát xét)."""
    try:
        cols = review_service.clean_custom_columns(body.get("columns"))
    except ValueError as e:
        raise HTTPException(400, str(e))
    with get_connection() as conn:
        if not conn.execute("SELECT 1 FROM mep_packages WHERE id=?", (pid,)).fetchone():
            raise HTTPException(404, "Không tìm thấy gói thầu")
        old = {c["id"]: c["label"] for c in review_service.custom_columns(conn, pid)}
        # Thứ tự: cột mới chưa có id → client gửi "new:<vị trí trong columns>"
        order = [("x_" + cols[int(k[4:])]["id"]) if isinstance(k, str) and k.startswith("new:") and k[4:].isdigit()
                 and int(k[4:]) < len(cols) else k for k in (body.get("order") or [])]
        order = review_service.normalize_order(order, cols)
        conn.execute("UPDATE mep_packages SET custom_cols_json=?, p2_order_json=?, updated_at=? WHERE id=?",
                     (json.dumps(cols, ensure_ascii=False), json.dumps(order), now_iso(), pid))
        if "round_unit" in body:  # làm tròn đơn giá: 0 / 10 / 100 / 1000
            ru = body.get("round_unit")
            if ru not in (0, 10, 100, 1000):
                raise HTTPException(400, "Làm tròn đơn giá chỉ nhận 0, 10, 100, 1000")
            conn.execute("UPDATE mep_packages SET round_unit=? WHERE id=?", (ru, pid))
        sid = (conn.execute("SELECT session_id FROM mep_packages WHERE id=?", (pid,)).fetchone() or {"session_id": None})["session_id"]
    added = [c["label"] for c in cols if c["id"] not in old]
    removed = [l for i, l in old.items() if i not in {c["id"] for c in cols}]
    if added or removed:
        review_service.log_event(pid, sid, "Cột soát xét: " + "; ".join(
            ([f"thêm {', '.join(added)}"] if added else []) + ([f"xóa {', '.join(removed)}"] if removed else [])))
    return {"columns": cols, "order": order}


@router.put("/packages/{pid}")
def update_package(pid: int, body: Dict[str, Any]):
    fields = ["project_id", "code", "name", "work_type", "subcontractor", "version_id", "deadline", "status", "vat_percent", "note", "he_so", "pa"]
    sets = {k: body[k] for k in fields if k in body}
    if "pa" in sets and sets["pa"] not in PA_LABELS:
        raise HTTPException(400, "Phương án không hợp lệ")
    if "he_so" in sets:
        try:
            sets["he_so"] = float(str(sets["he_so"]).replace(",", ".")) if sets["he_so"] not in (None, "") else 1.0
        except ValueError:
            raise HTTPException(400, "Hệ số phải là số")
        if sets["he_so"] <= 0:
            raise HTTPException(400, "Hệ số phải > 0")
    if "status" in sets and sets["status"] not in STATUSES:
        raise HTTPException(400, "Trạng thái không hợp lệ")
    if not sets:
        return {"updated": False}
    with get_connection() as conn:
        conn.execute(f"UPDATE mep_packages SET {', '.join(k + '=?' for k in sets)}, updated_at=? WHERE id=?",
                     (*sets.values(), now_iso(), pid))
    return {"updated": True}


@router.delete("/packages/{pid}")
def delete_package(pid: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM mep_packages WHERE id=?", (pid,))
    return {"deleted": True}


@router.post("/packages/{pid}/file")
async def replace_package_file(pid: int, file: UploadFile = File(...)):
    content = await file.read()
    try:
        wb_id = save_uploaded_workbook(content, file.filename or "bao-gia.xlsx")["id"]
    except ValueError as e:
        raise HTTPException(400, str(e))
    with get_connection() as conn:
        # Giữ session cũ để lần soát xét kế tiếp chép lại các dòng đã sửa tay (cùng sheet/dòng/hạng mục).
        conn.execute("UPDATE mep_packages SET workbook_id=?, sheet_name=NULL, updated_at=? WHERE id=?",
                     (wb_id, now_iso(), pid))
    return {"workbook_id": wb_id}


@router.post("/packages/{pid}/suggest-mapping")
def package_suggest_mapping(pid: int, body: Dict[str, Any]):
    with get_connection() as conn:
        row = conn.execute("SELECT workbook_id FROM mep_packages WHERE id=?", (pid,)).fetchone()
    if not row or not row["workbook_id"]:
        raise HTTPException(400, "Gói chưa có file báo giá NTP")
    return mep_router.suggest_boq_mapping({"workbook_id": row["workbook_id"], **body})


@router.post("/packages/{pid}/run")
def run_package_review(pid: int, body: Dict[str, Any]):
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM mep_packages WHERE id=?", (pid,)).fetchone()
    if not row:
        raise HTTPException(404, "Không tìm thấy gói thầu")
    p = dict(row)
    if not p.get("workbook_id"):
        raise HTTPException(400, "Gói chưa có file báo giá NTP")
    try:
        configs = review_sheets.resolve_configs(p, body)
        if (p.get("pa") or "PA2") == "PA1":
            _check_pa1(configs)
        res = review_sheets.run_all(p, configs, float(body.get("threshold_percent") or 10))
    except ValueError as e:
        raise HTTPException(400, str(e))
    # reset: về hồ sơ gốc – không giữ ô đã sửa, dòng đã thêm
    prev = None if body.get("reset") else p.get("session_id")
    kept = review_service.init_session(res["session_id"], prev)
    try:  # %KH soát xét theo thời gian (bảng KH vật tư theo vị trí dự án) cho các dòng chưa sửa tay
        with get_connection() as conn:
            region = _project_region(conn, p["project_id"])
        review_sheets.apply_kh({**p, "session_id": res["session_id"], "sheets_json": json.dumps(configs, ensure_ascii=False)}, region)
    except Exception:
        pass
    added = review_sheets.carry_added_rows(prev, res["session_id"])
    first = configs[0]
    with get_connection() as conn:
        conn.execute("UPDATE mep_packages SET session_id=?, sheet_name=?, header_row=?, mapping_json=?, sheets_json=?, status=CASE WHEN status IN ('Nháp','Đang soát xét') THEN 'Đang soát xét' ELSE status END, updated_at=? WHERE id=?",
                     (res["session_id"], first["sheet"], first["header_row"], json.dumps(first["mapping"]),
                      json.dumps(configs, ensure_ascii=False), now_iso(), pid))
    if (p.get("pa") or "PA2") == "PA1":  # Dự toán: tự áp giá E, F + %KH H
        try:
            with get_connection() as conn:
                region = _project_region(conn, p["project_id"])
            res["estimate_rows"] = _fill_estimate(pid, {**p, "session_id": res["session_id"],
                                                        "sheets_json": json.dumps(configs, ensure_ascii=False)}, region)
        except Exception as e:  # không chặn việc mở gói
            res["estimate_error"] = str(e)
    sheets_txt = ", ".join(f"{x['sheet']} ({x['total']})" for x in res["per_sheet"])
    _ = sheets_txt  # lịch sử chỉ ghi khi bấm Lưu
    res["kept_edits"] = kept["kept_edits"]
    res["sheet_name"] = first["sheet"]
    return res


def _reapply_kh_rows(pid: int, rows: List[Dict[str, Any]]):
    """Sửa ngày bắt đầu / kết thúc hoặc % ĐÃ KH trên bảng → %KH soát xét của dòng đó tính lại (trừ khi cùng lúc sửa tay %KH)."""
    ids = {int(r["id"]) for r in rows if r.get("id") and not str(r["id"]).startswith("new")
           and any(k in r for k in ("x_d_start", "x_d_end", "boq_dep_done")) and "rv_dep_pct" not in r}
    if not ids:
        return
    with get_connection() as conn:
        p = dict(conn.execute("SELECT * FROM mep_packages WHERE id=?", (pid,)).fetchone())
        infos = review_sheets.kh_infos(conn, p, _project_region(conn, p["project_id"]))
        for iid in ids:
            v = (infos.get(iid) or {}).get("pct")
            if v is not None:
                conn.execute("UPDATE mep_review_items SET rv_dep_pct=? WHERE id=?", (v, iid))
                review_sheets._set_dep_formula(conn, iid, (infos.get(iid) or {}).get("formula"))


@router.post("/packages/{pid}/save")
def save_package_rows(pid: int, body: Dict[str, Any], request: Request):
    """Lưu các ô soát xét người dùng sửa trên lưới (KL, VT, NC, Khấu hao, Đơn giá tổng, Ghi chú)."""
    user = unquote(request.headers.get("X-CCM-User", ""))[:80]
    with get_connection() as conn:
        p = conn.execute("SELECT session_id FROM mep_packages WHERE id=?", (pid,)).fetchone()
    if not p:
        raise HTTPException(404, "Không tìm thấy gói thầu")
    rows = body.get("rows") or []
    if rows and not p["session_id"]:
        raise HTTPException(400, "Gói chưa soát xét")
    try:
        res = review_service.save_rows(pid, p["session_id"], rows, user) if rows else {"saved": 0, "changes": 0}
    except ValueError as e:
        raise HTTPException(400, str(e))
    try:
        _reapply_kh_rows(pid, rows)
    except Exception:
        pass
    # Thao tác chưa có ô sửa (thêm / xóa dòng…) gửi kèm khi bấm Lưu → ghi lịch sử lúc này
    events = [ev.strip()[:300] for ev in (body.get("events") or [])[:50] if isinstance(ev, str) and ev.strip()]
    for ev in events:
        review_service.log_event(pid, p["session_id"], ev, user)
    # Mỗi lần Lưu có thay đổi → một bản lưu (số liệu + file Excel) để xem / tải lại sau
    if res.get("changes") or events:
        full = get_package(pid)
        v = review_versions.create(full, user, int(res.get("changes") or 0), "; ".join(events),
                                   lambda pk: review_export.export_review(pk, pk["items"], pk["sheet_configs"])["data"])
        res["version"] = v
    return res


@router.get("/packages/{pid}/versions")
def package_versions(pid: int):
    return review_versions.list_versions(pid)


@router.get("/packages/{pid}/versions/{vid}")
def package_version(pid: int, vid: int):
    v = review_versions.get(pid, vid)
    if not v:
        raise HTTPException(404, "Không tìm thấy bản lưu")
    v.pop("file_path", None)
    return v


@router.get("/packages/{pid}/versions/{vid}/file")
def package_version_file(pid: int, vid: int):
    v = review_versions.get(pid, vid)
    if not v or not v.get("file_path") or not Path(v["file_path"]).exists():
        raise HTTPException(404, "Bản lưu này không có file Excel")
    with get_connection() as conn:
        row = conn.execute("SELECT name FROM mep_packages WHERE id=?", (pid,)).fetchone()
    name = (row["name"] if row else "goi") + f" - lan luu {v['no']} ({v['saved_at'][:10]}).xlsx"
    return bytes_response(Path(v["file_path"]).read_bytes(), name,
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@router.get("/packages/{pid}/history")
def package_history(pid: int):
    return review_service.history(pid)


@router.get("/packages/{pid}/export-original")
def export_original(pid: int, request: Request):
    p = get_package(pid)
    if not p.get("session_id") or not p.get("workbook_id"):
        raise HTTPException(400, "Gói chưa soát xét")
    try:
        out = review_export.export_review(p, p["items"], p["sheet_configs"])
    except (ValueError, FileNotFoundError) as e:
        raise HTTPException(400, str(e))
    resp = bytes_response(out["data"], out["filename"], out["media_type"])
    resp.headers["X-CCM-Warning"] = "drawings" if out["drawings_warning"] else ""
    return resp


@router.get("/packages/{pid}/export")
def export_package(pid: int):
    """Bảng so sánh gọn (file mới), khác với export-original."""
    p = get_package(pid)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Soát xét giá"
    ws["A1"] = f"KẾT QUẢ SOÁT XÉT GIÁ – {p['code']} · {p['name']}"
    ws["A1"].font = Font(bold=True, size=14, color="1E3A8A")
    ws["A2"] = f"Dự án: {p['project_name']}   ·   Nhà thầu phụ: {p['subcontractor'] or '—'}   ·   Version giá: {p['version_name'] or '—'}"
    ws["A3"] = f"Giá NTP báo: {p['ntp_total'] or 0:,.0f} đ   ·   Giá soát xét: {p['review_total'] or 0:,.0f} đ   ·   Chênh lệch: {p['diff_percent'] if p['diff_percent'] is not None else '—'}%"
    rows = []
    for i, it in enumerate(p["items"], 1):
        rows.append([it.get("boq_stt") or i, it["boq_description"], it["boq_unit"],
                     it["boq_quantity"], it.get("boq_material"), it.get("boq_labor"), it.get("boq_machine"), it["ntp_total"], it["ntp_amount"], it.get("boq_note") or "",
                     it.get("rv_quantity"), it.get("rv_material"), it.get("rv_labor"), it.get("rv_machine"), it["rv_unit_total"], it["rv_amount"],
                     it["diff_amount"], it.get("lib_price"), it["review_status"], it.get("rv_note") or ""])
    _write_table(ws, 5, ["STT", "Hạng mục", "ĐVT", "KL (NTP)", "Đơn giá VT (NTP)", "Đơn giá NC (NTP)", "Khấu hao (NTP)",
                         "Đơn giá tổng (NTP)", "Thành tiền (NTP)", "Ghi chú NTP", "KL soát xét", "Đơn giá VT soát xét",
                         "Đơn giá NC soát xét", "Khấu hao soát xét", "Đơn giá tổng soát xét", "Thành tiền soát xét",
                         "Chênh lệch", "Giá thư viện", "Kết quả", "Ghi chú soát xét"],
                 rows, [7, 40, 7, 11, 14, 14, 13, 15, 16, 22, 11, 14, 14, 13, 15, 16, 15, 14, 12, 30])
    ws.freeze_panes = "C6"
    return _xlsx_response(wb, f"Soat-xet-{p['code']}.xlsx")


@router.post("/packages/{pid}/price-source")
def package_price_source(pid: int, body: Dict[str, Any]):
    """Giá theo nguồn (nút Chạy): default / dntc / ncc / hd / truc → đơn giá; kh → %KH theo thời gian.
    Chỉ trả số liệu, giao diện ghi vào các ô đang chọn (chưa lưu cho tới khi bấm Lưu)."""
    from backend import price_sources
    source = str(body.get("source") or "default")
    ids = {int(x) for x in (body.get("ids") or [])}
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM mep_packages WHERE id=?", (pid,)).fetchone()
        if not row or not row["session_id"]:
            raise HTTPException(400, "Gói chưa soát xét")
        p = dict(row)
        if source == "kh":
            infos = review_sheets.kh_infos(conn, p, _project_region(conn, p["project_id"]))
            dep = {str(k): v["pct"] for k, v in infos.items() if v.get("pct") is not None and (not ids or k in ids)}
            fml = {str(k): v["formula"] for k, v in infos.items() if v.get("formula") and (not ids or k in ids)}
            from backend import depreciation as dp
            t = dp.load_tables(conn)
            return {"dep": dep, "formula": fml, "used": [f"Khấu hao {t['version']}" if t.get("version") else "chưa nạp bảng Khấu hao"]}
        items = [i for i in review_service.session_items(conn, p["session_id"]) if not ids or i["id"] in ids]
        s = conn.execute("SELECT threshold_percent FROM mep_review_sessions WHERE id=?", (p["session_id"],)).fetchone()
        res = price_sources.prices(conn, items, source, float((s["threshold_percent"] if s else None) or 10),
                                   int(body["version_id"]) if body.get("version_id") else None)
        hs = float(p.get("he_so") or 1)
        if source in ("truc", "ncc") and hs != 1:  # nhân công trực / giá báo giá NCC × hệ số của gói
            for x in res["prices"].values():
                for k in ("material", "labor", "machine", "total"):
                    if x.get(k) is not None:
                        x[k] = float(x[k]) * hs
            res["used"].append(f"× hệ số {hs:g}")
        return res
