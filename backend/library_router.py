"""API Thư viện giá (version, nạp file, sửa trực tiếp, so sánh, xuất Excel)."""
import json
import uuid
from typing import Any, Dict, Optional
from urllib.parse import unquote

import openpyxl
from openpyxl.styles import Font, PatternFill
from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile

from backend import library_service as lib, library_structure
from backend.mep_db import get_connection
from backend.mep_excel import UPLOADS_DIR
from backend.xlsx_utils import bytes_response, write_table, xlsx_response

router = APIRouter(prefix="/api/mep/library", tags=["Thư viện giá"])
library_structure.ensure_schema()

MAX_UPLOAD = 25 * 1024 * 1024
TEMPLATE_HEADERS = ["Loại", "Mã", "Tên vật tư / công việc", "Quy cách / Hãng", "ĐVT", "Đơn giá VT", "Đơn giá NC",
                    "Đơn giá máy", "Hao hụt (%)", "Đơn giá tổng", "Nhóm"]
WIDTHS = [11, 12, 36, 24, 8, 14, 14, 14, 11, 15, 18]


def user_of(request: Request) -> str:
    return unquote(request.headers.get("X-CCM-User", ""))[:80]


def _fail(e: Exception):
    if isinstance(e, LookupError):
        raise HTTPException(404, str(e).strip("'\""))
    if isinstance(e, ValueError):
        raise HTTPException(400, str(e))
    raise e


async def _read_upload(file: UploadFile) -> bytes:
    name = (file.filename or "").lower()
    if not name.endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Chỉ nhận file Excel .xlsx hoặc .xlsm (file .xls cần lưu lại dạng .xlsx).")
    data = await file.read()
    if not data:
        raise HTTPException(400, "File rỗng")
    if len(data) > MAX_UPLOAD:
        raise HTTPException(413, "File vượt quá 25 MB")
    return data


def _parse_bytes(data: bytes, raw_only: bool = False) -> Dict[str, Any]:
    tmp = UPLOADS_DIR / f"lib_parse_{uuid.uuid4().hex}.xlsx"
    tmp.write_bytes(data)
    try:
        return lib.parse_library_file(tmp, raw_only=raw_only)
    finally:
        tmp.unlink(missing_ok=True)


def _row(r: Dict[str, Any]):
    return [r.get("item_type") or "Vật tư", r["code"], r["description"], r["specification"], r["unit"],
            r["material_price"], r["labor_price"], r.get("machine_price"), r.get("waste_percent"), r["total_price"],
            r.get("category") or r.get("system")]


@router.get("/template")
def template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Thu vien gia"
    write_table(ws, 1, TEMPLATE_HEADERS, [
        ["Vật tư", "CD-0012", "Cáp điện CV 3×4mm²", "Cadivi, 0,6/1kV", "m", 18500, 6500, None, 3, 25000, "Cáp điện"],
        ["Vật tư", "NU-0402", "Ống PPR D25 cấp nước", "Bình Minh PN20", "m", 42000, 18000, None, 5, 60000, "Cấp thoát nước"],
        ["Nhân công", "NC-01", "Nhân công điện bậc 3,5/7", "", "công", None, 450000, None, None, 450000, "Nhân công"],
        ["Ca máy", "CM-01", "Máy khoan bê tông 1,5kW", "", "ca", None, None, 285000, None, 285000, "Ca máy"],
    ], WIDTHS)
    ws.freeze_panes = "A2"
    note = wb.create_sheet("Huong dan")
    for i, line in enumerate([
        "Mỗi dòng là một mục giá. Bắt buộc: Tên, ĐVT và ít nhất một cột đơn giá.",
        "Loại: Vật tư / Nhân công / Ca máy. Bỏ trống thì tool tự đoán theo tên sheet, mã (NC-, CM-) hoặc ĐVT (công, ca).",
        "Đơn giá tổng bỏ trống thì tool tự cộng VT + NC + Máy.",
        "Có thể tách Vật tư, Nhân công, Ca máy ra các sheet riêng; tool đọc mọi sheet có tiêu đề cột phù hợp.",
        "Dòng thiếu tên, thiếu ĐVT, thiếu giá, giá âm hoặc trùng mã sẽ bị bỏ qua và liệt kê ở mục 'dòng lỗi'.",
    ], 1):
        note.cell(row=i, column=1, value=line)
    note.column_dimensions["A"].width = 110
    return xlsx_response(wb, "Mau-thu-vien-gia.xlsx")


LIB_KINDS = {"dntc": "Đơn giá ĐNTC chuẩn", "ncc": "Báo giá NCC", "hd": "Hợp đồng tham khảo", "kh": "Khấu hao",
             "truc": "Đơn giá nhân công trực"}


@router.get("/versions")
def versions(include_reference: bool = False):
    """Version thư viện. File đối chiếu riêng của gói (status='reference') chỉ trả khi include_reference=true."""
    with get_connection() as conn:
        rows = conn.execute("SELECT * FROM mep_library_versions WHERE ? OR status<>'reference'"
                            " ORDER BY COALESCE(effective_date,'') DESC, id DESC", (1 if include_reference else 0,)).fetchall()
        return [dict(r) for r in rows]


@router.post("/versions/parse")
async def parse_file(file: UploadFile = File(...), compare_to: Optional[int] = Form(None)):
    data = await _read_upload(file)
    try:
        parsed = _parse_bytes(data)
    except Exception as e:
        _fail(e)
    with get_connection() as conn:
        base = compare_to or lib.active_version_id(conn)
    parsed.pop("structure", None)
    return {"filename": file.filename, **parsed, **lib.preview_against(parsed["entries"], base)}


@router.post("/versions/import")
async def import_file(request: Request, file: UploadFile = File(...), activate: Optional[bool] = Form(None), name: str = Form(""),
                      kind: str = Form("dntc")):
    """Nạp file → kiểm tra → lưu thành version mới. Đơn giá ĐNTC chuẩn: áp dụng ngay; Báo giá NCC / Hợp đồng tham khảo: chỉ lưu."""
    kind = kind if kind in LIB_KINDS else "dntc"
    if activate is None:
        activate = kind == "dntc"
    data = await _read_upload(file)
    try:
        # Khấu hao / Đơn giá nhân công trực: giữ nguyên mọi sheet như file (bảng tra cứu, không tách mục giá)
        parsed = _parse_bytes(data, raw_only=kind in ("kh", "truc"))
        with get_connection() as conn:
            base = lib.active_version_id(conn)
        preview = lib.preview_against(parsed["entries"], base) if kind not in ("kh", "truc") else {"changed": [], "new": [], "removed": []}
        created = lib.create_version(name, parsed["entries"], note=f"Nạp từ {file.filename}", created_by=user_of(request),
                                     error_count=len(parsed["errors"]), structure=parsed["structure"],
                                     source_bytes=data, source_name=file.filename or "")
        with get_connection() as conn:
            conn.execute("UPDATE mep_library_versions SET kind=? WHERE id=?", (kind, created["id"]))
        if activate:
            lib.activate_version(created["id"])
    except Exception as e:
        _fail(e)
    return {"version": created, "kind": kind, "activated": activate, "filename": file.filename, "row_count": parsed["row_count"],
            "valid_count": parsed["valid_count"], "by_type": parsed["by_type"], "sheets": parsed["sheets"],
            "skipped_sheets": parsed["skipped_sheets"],
            "errors": parsed["errors"], "warnings": parsed["warnings"],
            "changed": preview["changed"], "new": preview["new"], "removed": preview["removed"]}


@router.post("/versions")
def create(body: Dict[str, Any], request: Request):
    try:
        return lib.create_version(body.get("name", ""), body.get("entries") or [], body.get("effective_date", ""),
                                  body.get("note", ""), body.get("created_by") or user_of(request),
                                  "draft" if body.get("draft") else "checked", int(body.get("error_count") or 0),
                                  body.get("scope") or "Toàn bộ (Điện + Nước)")
    except Exception as e:
        _fail(e)


@router.put("/versions/{vid}")
def rename(vid: int, body: Dict[str, Any]):
    sets = {k: str(body[k]).strip() for k in ("name", "effective_date", "note", "kind") if k in body}
    if "kind" in sets and sets["kind"] not in LIB_KINDS:
        raise HTTPException(400, "Phân loại không hợp lệ")
    if "name" in sets and not sets["name"]:
        raise HTTPException(400, "Tên version không được trống")
    with get_connection() as conn:
        if "name" in sets and conn.execute("SELECT 1 FROM mep_library_versions WHERE name=? AND id<>?", (sets["name"], vid)).fetchone():
            raise HTTPException(400, "Tên version đã tồn tại")
        if sets:
            conn.execute(f"UPDATE mep_library_versions SET {', '.join(k + '=?' for k in sets)} WHERE id=?", (*sets.values(), vid))
    return {"updated": bool(sets)}


@router.post("/versions/{vid}/activate")
def activate(vid: int):
    with get_connection() as conn:
        k = conn.execute("SELECT kind FROM mep_library_versions WHERE id=?", (vid,)).fetchone()
    if k and k[0] in ("kh", "truc"):
        raise HTTPException(400, "Bảng Khấu hao / Nhân công trực chỉ để tra cứu, không đặt làm thư viện giá hiện hành")
    try:
        lib.activate_version(vid)
    except Exception as e:
        _fail(e)
    return {"active": vid}


@router.delete("/versions/{vid}")
def delete_version(vid: int):
    with get_connection() as conn:
        v = conn.execute("SELECT status FROM mep_library_versions WHERE id=?", (vid,)).fetchone()
        if not v:
            raise HTTPException(404, "Không tìm thấy version")
        # Version khác còn lại (mới nhất) để thay thế: đặt làm hiện hành nếu xóa bản đang áp dụng,
        # và chuyển các gói thầu đang dùng version bị xóa sang bản đó (số liệu đã soát xét của gói giữ nguyên)
        repl = conn.execute("""SELECT id FROM mep_library_versions WHERE id<>? AND status<>'reference'
                               ORDER BY (status='active') DESC, COALESCE(effective_date,'') DESC, id DESC LIMIT 1""", (vid,)).fetchone()
        repl_id = repl[0] if repl else None
        used = conn.execute("SELECT COUNT(*) FROM mep_packages WHERE version_id=?", (vid,)).fetchone()[0]
        if used and repl_id is None:
            raise HTTPException(400, "Đây là version duy nhất và đang được gói thầu sử dụng – nạp version khác trước rồi mới xóa")
        if used:
            conn.execute("UPDATE mep_packages SET version_id=? WHERE version_id=?", (repl_id, vid))
        if v["status"] == "active" and repl_id is not None:
            conn.execute("UPDATE mep_library_versions SET status='active' WHERE id=?", (repl_id,))
        conn.execute("DELETE FROM mep_libraries WHERE version=?", (vid,))
        conn.execute("DELETE FROM mep_library_edits WHERE version_id=?", (vid,))
        conn.execute("DELETE FROM mep_library_rows WHERE version_id=?", (vid,))
        conn.execute("DELETE FROM mep_library_sheets WHERE version_id=?", (vid,))
        src = conn.execute("SELECT source_file FROM mep_library_versions WHERE id=?", (vid,)).fetchone()[0]
        if src:
            from pathlib import Path
            Path(src).unlink(missing_ok=True)
        conn.execute("DELETE FROM mep_library_versions WHERE id=?", (vid,))
    return {"deleted": True, "active": repl_id if v["status"] == "active" else None, "moved_packages": used}


@router.get("/versions/{vid}/items")
def items(vid: int, type: str = "", category: str = "", q: str = "", compare_to: Optional[int] = None,
          sheet_id: Optional[int] = None):
    try:
        return lib.list_items(vid, type, category, q, compare_to, sheet_id)
    except Exception as e:
        _fail(e)


@router.post("/versions/{vid}/items")
def add_item(vid: int, body: Dict[str, Any], request: Request):
    try:
        return lib.add_item(vid, body, user_of(request), body.get("sheet_id"))
    except Exception as e:
        _fail(e)


@router.put("/items/{item_id}")
def update_item(item_id: int, body: Dict[str, Any], request: Request):
    try:
        return lib.update_item(item_id, body, user_of(request))
    except Exception as e:
        _fail(e)


@router.delete("/items/{item_id}")
def delete_item(item_id: int, request: Request):
    try:
        lib.delete_item(item_id, user_of(request))
    except Exception as e:
        _fail(e)
    return {"deleted": True}


@router.get("/versions/{vid}/history")
def version_history(vid: int):
    return lib.history(vid)


@router.get("/compare")
def compare(base: int, target: int):
    try:
        return lib.compare_versions(base, target)
    except Exception as e:
        _fail(e)


@router.get("/compare/export")
def compare_export(base: int, target: int):
    try:
        res = lib.compare_versions(base, target)
    except Exception as e:
        _fail(e)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tong hop"
    ws["A1"] = f"SO SÁNH THƯ VIỆN GIÁ: {res['base']['name']} → {res['target']['name']}"
    ws["A1"].font = Font(bold=True, size=14, color="1E3A8A")
    for i, (k, label) in enumerate([("added", "Thêm mới"), ("changed", "Thay đổi"), ("removed", "Bị xóa"), ("unchanged", "Không đổi")], 3):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=res["summary"][k])
    ws.column_dimensions["A"].width = 16
    heads = ["Loại", "Mã", "Tên", "ĐVT", "Đơn giá VT", "Đơn giá NC", "Đơn giá máy", "Đơn giá tổng"]
    simple = lambda r: [r["item_type"], r["code"], r["description"], r["unit"], r["material_price"], r["labor_price"], r["machine_price"], r["total_price"]]
    sh = wb.create_sheet("Them moi")
    write_table(sh, 1, heads, [simple(r) for r in res["added"]], [11, 12, 40, 8, 14, 14, 14, 15])
    sh = wb.create_sheet("Bi xoa")
    write_table(sh, 1, heads, [simple(r) for r in res["removed"]], [11, 12, 40, 8, 14, 14, 14, 15])
    sh = wb.create_sheet("Thay doi")
    write_table(sh, 1, ["Loại", "Mã", "Tên", "ĐVT", "Trường thay đổi", f"Đơn giá tổng {res['base']['name']}",
                        f"Đơn giá tổng {res['target']['name']}", "Chênh lệch (%)"],
                [[r["item_type"], r["code"], r["description"], r["unit"], ", ".join(r["field_labels"]), r["old_total"],
                  r["total_price"], r["diff_percent"]] for r in res["changed"]], [11, 12, 40, 8, 26, 18, 18, 14])
    red, green = PatternFill("solid", start_color="FDE8E8"), PatternFill("solid", start_color="E7F6EC")
    for row in sh.iter_rows(min_row=2, min_col=8, max_col=8):
        for c in row:
            if isinstance(c.value, (int, float)):
                c.fill = red if c.value > 0 else green if c.value < 0 else PatternFill()
    return xlsx_response(wb, f"So-sanh-{res['base']['name']}-{res['target']['name']}.xlsx")


@router.get("/versions/{vid}/export")
def export_version(vid: int, type: str = "", category: str = "", q: str = ""):
    with get_connection() as conn:
        v = conn.execute("SELECT * FROM mep_library_versions WHERE id=?", (vid,)).fetchone()
    if not v:
        raise HTTPException(404, "Không tìm thấy version")
    if not (type or category or q):
        # Version nạp từ file: trả lại đúng file gốc (đủ sheet, cột, định dạng) kèm các ô đã sửa.
        out = library_structure.export_original(vid)
        if out:
            media = "application/vnd.ms-excel.sheet.macroEnabled.12" if out["ext"] == ".xlsm" else \
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            return bytes_response(out["data"], f"Thu-vien-gia-{v['name']}{out['ext']}", media)
    data = lib.list_items(vid, type, category, q)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Thu vien gia"
    write_table(ws, 1, TEMPLATE_HEADERS, [_row(r) for r in data["items"]], WIDTHS)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    info = wb.create_sheet("Thong tin")
    for i, (k, val) in enumerate([("Version", v["name"]), ("Hiệu lực", v["effective_date"]), ("Trạng thái", v["status"]),
                                  ("Số mục", len(data["items"])), ("Ghi chú", v["note"]),
                                  ("Bộ lọc", ", ".join(x for x in (type, category, q) if x) or "Không")], 1):
        info.cell(row=i, column=1, value=k).font = Font(bold=True)
        info.cell(row=i, column=2, value=val)
    info.column_dimensions["A"].width, info.column_dimensions["B"].width = 14, 60
    return xlsx_response(wb, f"Thu-vien-gia-{v['name']}.xlsx")


# ------------------------------------------------------------------ vật tư khấu hao (đã học)

@router.get("/depreciation")
def depreciation_summary():
    from backend import depreciation as dp
    return dp.summary()


@router.post("/depreciation/learn")
async def depreciation_learn(file: UploadFile = File(...)):
    """Học từ hồ sơ BOQ có cột % KH VẬT TƯ: vật tư nào có thể khấu hao < 100%, vật tư nào mặc định 100%."""
    from backend import depreciation as dp
    data = await _read_upload(file)
    tmp = UPLOADS_DIR / f"kh_learn_{uuid.uuid4().hex}.xlsx"
    tmp.write_bytes(data)
    try:
        r = dp.learn_from_boq(tmp, file.filename or "")
    except Exception as e:
        _fail(e)
    finally:
        tmp.unlink(missing_ok=True)
    if not r["rows"]:
        raise HTTPException(400, "Không thấy bảng có cột Mô tả + % KH VẬT TƯ trong file")
    return r


@router.post("/depreciation/set")
def depreciation_set(body: Dict[str, Any]):
    """Đổi loại 1 tên: kh = 'lt100' (có thể < 100%) / '100' (mặc định 100%)."""
    from backend import depreciation as dp
    name, kh = str(body.get("name") or "").strip(), body.get("kh")
    if not name or kh not in ("lt100", "100"):
        raise HTTPException(400, "Thiếu tên hoặc loại")
    data = dp._load(dp.LEARN_FILE)
    data[dp.norm(name)] = {"name": name, "kh": kh, "src": "sửa tay"}
    dp.LEARN_FILE.parent.mkdir(parents=True, exist_ok=True)
    dp.LEARN_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    dp._cache.clear()
    return dp.summary()

