"""FastAPI Router for CCM-AI MEP Price Review System."""
import json
import uuid
import openpyxl
from pathlib import Path
from typing import Dict, Any, List, Optional
from fastapi import APIRouter, HTTPException, UploadFile, File, Form, Query
from fastapi.responses import FileResponse

from backend.mep_db import get_connection, now_iso
from backend import boq_reader
from backend.xlsx_utils import parse_number
ROOT = Path(__file__).resolve().parents[1]
from backend.mep_engine import (
    normalize_mep_description,
    is_non_item_row,
    match_single_item,
    extract_mep_specs
)
from backend.mep_excel import (
    save_uploaded_workbook,
    get_workbook_record,
    get_sheet_grid,
    update_sheet_cells,
    apply_review_highlights,
    generate_variance_report,
    UPLOADS_DIR,
    REPORTS_DIR
)

router = APIRouter(prefix="/api/mep", tags=["MEP Price Review"])


def _num(value) -> Optional[float]:
    """Parses a numeric Excel cell; returns None for blanks/text."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def dep_fraction(value) -> Optional[float]:
    """% khấu hao: 0,2 / 20 / '20%' → 0.2 (phân số)."""
    n = parse_number(value)
    if n is None:
        return None
    return n / 100 if n > 1 else n


def _ensure_review_columns():
    """Adds columns used by the package review grid (idempotent)."""
    with get_connection() as conn:
        cols = {r[1] for r in conn.execute("PRAGMA table_info(mep_review_items)")}
        for name, typ in (("boq_stt", "TEXT"), ("boq_material", "REAL"), ("boq_labor", "REAL"),
                          ("boq_machine", "REAL"), ("boq_amount", "REAL"), ("boq_note", "TEXT"),
                          ("rv_quantity", "REAL"), ("rv_material", "REAL"), ("rv_labor", "REAL"),
                          ("rv_machine", "REAL"), ("rv_total", "REAL"), ("rv_note", "TEXT"),
                          ("rv_touched", "INTEGER DEFAULT 0"), ("boq_dep_pct", "REAL"), ("boq_dep_done", "REAL"), ("rv_dep_pct", "REAL"), ("rv_extra", "TEXT"), ("is_added", "INTEGER DEFAULT 0")):
            if name not in cols:
                conn.execute(f"ALTER TABLE mep_review_items ADD COLUMN {name} {typ}")


_ensure_review_columns()

# ==================== 1. WORKBOOK / EXCEL WORKSPACE ====================

@router.post("/workbooks/upload")
async def upload_boq_workbook(file: UploadFile = File(...)):
    """Uploads BOQ Excel file (.xlsx or .xlsm), returns sheet list and metadata."""
    filename = file.filename or "boq.xlsx"
    ext = Path(filename).suffix.lower()
    if ext not in ['.xlsx', '.xlsm']:
        raise HTTPException(400, "Chỉ chấp nhận file Excel .xlsx hoặc .xlsm")
    
    content = await file.read()
    if len(content) > 50 * 1024 * 1024:
        raise HTTPException(413, "File vượt quá 50MB")

    try:
        res = save_uploaded_workbook(content, filename)
        return res
    except Exception as e:
        raise HTTPException(500, f"Lỗi xử lý file Excel: {str(e)}")

@router.get("/workbooks")
def list_workbooks():
    """Lists all uploaded workbooks."""
    with get_connection() as conn:
        rows = conn.execute("SELECT id, filename, original_filename, file_type, sheets_info, uploaded_at FROM mep_workbooks ORDER BY uploaded_at DESC").fetchall()
        result = []
        for r in rows:
            d = dict(r)
            if d.get("sheets_info"):
                try: d["sheets"] = json.loads(d["sheets_info"])
                except: d["sheets"] = []
            result.append(d)
        return result

@router.get("/workbooks/sample")
def get_or_create_sample_workbook():
    """Returns or loads sample BOQ workbook."""
    sample_path = ROOT / "BoQ-MEP-Mau-Can-Soat-Gia.xlsx"
    if not sample_path.exists():
        raise HTTPException(404, "Không tìm thấy file mẫu")
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM mep_workbooks WHERE original_filename = 'BoQ-MEP-Mau-Can-Soat-Gia.xlsx' ORDER BY uploaded_at DESC LIMIT 1").fetchone()
        if row:
            rec = dict(row)
            if rec.get("sheets_info"):
                try: rec["sheets"] = json.loads(rec["sheets_info"])
                except: rec["sheets"] = []
            return rec
    content = sample_path.read_bytes()
    return save_uploaded_workbook(content, "BoQ-MEP-Mau-Can-Soat-Gia.xlsx")

@router.get("/workbooks/{wb_id}/sheet")
def get_workbook_sheet(wb_id: str, sheet_name: Optional[str] = None):
    """Retrieves sheet cells, formats, merges, and active review warning highlights."""
    try:
        return get_sheet_grid(wb_id, sheet_name)
    except Exception as e:
        raise HTTPException(500, f"Lỗi đọc sheet Excel: {str(e)}")

@router.post("/workbooks/{wb_id}/cells")
def update_workbook_cells(wb_id: str, body: Dict[str, Any]):
    """Saves cell edits from Excel Web Editor."""
    sheet_name = body.get("sheet_name")
    updates = body.get("updates", [])
    if not sheet_name or not isinstance(updates, list):
        raise HTTPException(400, "Thiếu sheet_name hoặc danh sách updates")

    try:
        return update_sheet_cells(wb_id, sheet_name, updates)
    except Exception as e:
        raise HTTPException(500, f"Lỗi cập nhật ô Excel: {str(e)}")

@router.post("/workbooks/{wb_id}/structure")
def update_workbook_structure(wb_id: str, body: Dict[str, Any]):
    """Modifies sheet structure: insert/delete rows/columns or add/delete/rename sheets."""
    from backend.mep_excel import modify_sheet_structure
    sheet_name = body.get("sheet_name", "Sheet1")
    action = body.get("action")
    index = int(body.get("index") or 1)
    extra_name = body.get("extra_name")
    if not action:
        raise HTTPException(400, "Thiếu tham số action")
    try:
        return modify_sheet_structure(wb_id, sheet_name, action, index, extra_name)
    except Exception as e:
        raise HTTPException(500, f"Lỗi thao tác cấu trúc Excel: {str(e)}")

@router.get("/workbooks/{wb_id}/download")
def download_workbook(wb_id: str):
    """Downloads current active workbook."""
    rec = get_workbook_record(wb_id)
    path = Path(rec["file_path"])
    if not path.is_file():
        raise HTTPException(404, "File không tồn tại")
    return FileResponse(path, filename=rec["original_filename"])

@router.get("/workbooks/{wb_id}/download-reviewed")
def download_reviewed_boq(wb_id: str, session_id: str):
    """Downloads reviewed BOQ with highlighted cells and inspection comments."""
    try:
        reviewed_path = apply_review_highlights(wb_id, session_id)
        rec = get_workbook_record(wb_id)
        name_part = Path(rec["original_filename"]).stem
        ext = rec["file_type"]
        out_name = f"{name_part}_SoatGia{ext}"
        return FileResponse(reviewed_path, filename=out_name)
    except Exception as e:
        raise HTTPException(500, f"Lỗi tạo file BOQ sau soát xét: {str(e)}")

# ==================== 2. MEP PRICE LIBRARY ====================

@router.post("/libraries/upload")
async def upload_library_file(file: UploadFile = File(...)):
    """Uploads a price library file, returns sheets and header preview for mapping."""
    content = await file.read()
    filename = file.filename or "library.xlsx"
    temp_path = UPLOADS_DIR / f"lib_temp_{uuid.uuid4().hex}.xlsx"
    temp_path.write_bytes(content)

    try:
        wb = openpyxl.load_workbook(temp_path, read_only=True, data_only=True)
        sheets = []
        for sname in wb.sheetnames:
            ws = wb[sname]
            first_rows = []
            for row in ws.iter_rows(max_row=5, values_only=True):
                first_rows.append([str(v) if v is not None else "" for v in row])
            sheets.append({
                "name": sname,
                "preview": first_rows
            })
        wb.close()
        temp_path.unlink(missing_ok=True)
        return {"filename": filename, "sheets": sheets}
    except Exception as e:
        temp_path.unlink(missing_ok=True)
        raise HTTPException(500, f"Lỗi đọc file thư viện: {str(e)}")

@router.post("/libraries/commit")
def commit_library_entries(body: Dict[str, Any]):
    """Saves price library entries into SQLite with MEP system tag."""
    system = body.get("system", "Khác")
    source = body.get("source", "Thư viện Excel")
    project = body.get("project", "")
    vendor = body.get("vendor", "")
    date_val = body.get("date", now_iso()[:10])
    entries = body.get("entries", [])

    if not isinstance(entries, list) or not entries:
        raise HTTPException(400, "Danh sách entries không hợp lệ")

    inserted = 0
    with get_connection() as conn:
        for it in entries:
            desc = str(it.get("description") or "").strip()
            price = float(it.get("total_price") or it.get("material_price") or 0.0)
            if not desc or price <= 0:
                continue

            combined = f"{desc} {it.get('specification','')} {it.get('brand','')} {it.get('model','')}"
            norm_desc = normalize_mep_description(combined)

            conn.execute("""
                INSERT INTO mep_libraries (
                    system, code, description, specification, brand, origin, model, unit,
                    material_price, labor_price, total_price, source, vendor, project, date,
                    version, normalized_desc, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                system,
                str(it.get("code") or "").strip(),
                desc,
                str(it.get("specification") or "").strip(),
                str(it.get("brand") or "").strip(),
                str(it.get("origin") or "").strip(),
                str(it.get("model") or "").strip(),
                str(it.get("unit") or "").strip(),
                it.get("material_price"),
                it.get("labor_price"),
                price,
                source,
                vendor,
                project,
                date_val,
                1,
                norm_desc,
                now_iso()
            ))
            inserted += 1

    return {"saved": True, "count": inserted, "system": system}

@router.get("/libraries")
def get_library_items(
    system: Optional[str] = None,
    query: Optional[str] = None,
    limit: int = 200,
    offset: int = 0
):
    """Searches and lists price library entries."""
    with get_connection() as conn:
        sql = "SELECT * FROM mep_libraries WHERE 1=1"
        params = []
        if system and system != "Tất cả":
            sql += " AND system = ?"
            params.append(system)
        if query and query.strip():
            sql += " AND (description LIKE ? OR code LIKE ? OR specification LIKE ? OR brand LIKE ? OR model LIKE ?)"
            q = f"%{query.strip()}%"
            params.extend([q, q, q, q, q])

        count_sql = "SELECT COUNT(*) FROM (" + sql + ")"
        total = conn.execute(count_sql, params).fetchone()[0]

        sql += " ORDER BY id DESC LIMIT ? OFFSET ?"
        params.extend([limit, offset])

        rows = conn.execute(sql, params).fetchall()
        return {
            "total": total,
            "items": [dict(r) for r in rows]
        }

@router.post("/libraries/item")
def add_library_item(item: Dict[str, Any]):
    """Adds a single library item manually."""
    desc = str(item.get("description") or "").strip()
    price = float(item.get("total_price") or item.get("material_price") or 0.0)
    if not desc or price <= 0:
        raise HTTPException(400, "Cần mô tả và đơn giá > 0")

    combined = f"{desc} {item.get('specification','')} {item.get('brand','')} {item.get('model','')}"
    norm_desc = normalize_mep_description(combined)

    with get_connection() as conn:
        cur = conn.execute("""
            INSERT INTO mep_libraries (
                system, code, description, specification, brand, origin, model, unit,
                material_price, labor_price, total_price, source, vendor, project, date,
                version, normalized_desc, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            item.get("system", "Khác"),
            str(item.get("code") or "").strip(),
            desc,
            str(item.get("specification") or "").strip(),
            str(item.get("brand") or "").strip(),
            str(item.get("origin") or "").strip(),
            str(item.get("model") or "").strip(),
            str(item.get("unit") or "").strip(),
            item.get("material_price"),
            item.get("labor_price"),
            price,
            item.get("source", "Thêm thủ công"),
            item.get("vendor", ""),
            item.get("project", ""),
            item.get("date", now_iso()[:10]),
            1,
            norm_desc,
            now_iso()
        ))
        new_id = cur.lastrowid
    return {"id": new_id, "saved": True}

@router.put("/libraries/item/{item_id}")
def update_library_item(item_id: int, body: Dict[str, Any]):
    """Updates library item and records price history."""
    with get_connection() as conn:
        old = conn.execute("SELECT * FROM mep_libraries WHERE id = ?", (item_id,)).fetchone()
        if not old:
            raise HTTPException(404, "Không tìm thấy vật tư")

        new_price = float(body.get("total_price") or body.get("material_price") or old["total_price"])
        if new_price != old["total_price"]:
            conn.execute("""
                INSERT INTO mep_price_history (library_id, old_price, new_price, changed_by, reason, timestamp)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (item_id, old["total_price"], new_price, "Người dùng", body.get("reason", "Cập nhật đơn giá"), now_iso()))

        desc = str(body.get("description") or old["description"]).strip()
        combined = f"{desc} {body.get('specification', old['specification'])} {body.get('brand', old['brand'])} {body.get('model', old['model'])}"
        norm_desc = normalize_mep_description(combined)

        conn.execute("""
            UPDATE mep_libraries SET
                system = ?, code = ?, description = ?, specification = ?, brand = ?,
                origin = ?, model = ?, unit = ?, material_price = ?, labor_price = ?,
                total_price = ?, source = ?, vendor = ?, project = ?, date = ?,
                version = version + 1, normalized_desc = ?
            WHERE id = ?
        """, (
            body.get("system", old["system"]),
            body.get("code", old["code"]),
            desc,
            body.get("specification", old["specification"]),
            body.get("brand", old["brand"]),
            body.get("origin", old["origin"]),
            body.get("model", old["model"]),
            body.get("unit", old["unit"]),
            body.get("material_price", old["material_price"]),
            body.get("labor_price", old["labor_price"]),
            new_price,
            body.get("source", old["source"]),
            body.get("vendor", old["vendor"]),
            body.get("project", old["project"]),
            body.get("date", old["date"]),
            norm_desc,
            item_id
        ))

    return {"updated": True, "id": item_id}

@router.delete("/libraries/item/{item_id}")
def delete_library_item(item_id: int):
    """Deletes a library item."""
    with get_connection() as conn:
        conn.execute("DELETE FROM mep_libraries WHERE id = ?", (item_id,))
    return {"deleted": True, "id": item_id}

# ==================== 3. BOQ MAPPING & REVIEW ENGINE ====================

@router.post("/review/suggest-mapping")
def suggest_boq_mapping(body: Dict[str, Any]):
    """Auto-detects header row and proposes column mapping for 9 BOQ fields."""
    wb_id = body.get("workbook_id")
    sheet_name = body.get("sheet_name")
    header_row_candidate = int(body.get("header_row") or 1)

    rec = get_workbook_record(wb_id)
    wb = openpyxl.load_workbook(Path(rec["file_path"]), data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            sheet_name = wb.sheetnames[0]
        sheet = wb[sheet_name]
        if body.get("force_header_row"):
            # Người dùng chỉ định dòng tiêu đề: ghép với tiêu đề tầng trên (nếu có ô gộp).
            hr = max(1, header_row_candidate)
            child = boq_reader._row_texts(sheet, hr, min(sheet.max_column or 1, 40))
            parent = boq_reader._row_texts(sheet, hr - 1, min(sheet.max_column or 1, 40)) if hr > 1 else {}
            headers = {c: (f"{parent[c]} {t}" if parent.get(c) and boq_reader._norm(parent[c]) != boq_reader._norm(t)
                           and boq_reader._norm(t) in boq_reader.SUB_HINT else t) for c, t in child.items()}
            det = {"header_row": hr, "headers": headers, "mapping": boq_reader.map_columns(headers)}
        else:
            det = boq_reader.detect_header(sheet)
    finally:
        wb.close()
    return {"header_row": det["header_row"], "mapping": det["mapping"], "headers": det["headers"], "sheet_name": sheet_name}

@router.post("/review/run")
def run_price_review(body: Dict[str, Any]):
    """Executes deterministic MEP Price Review matching."""
    wb_id = body.get("workbook_id")
    sheet_name = body.get("sheet_name")
    header_row = int(body.get("header_row") or 1)
    mapping = body.get("mapping", {})
    threshold_percent = float(body.get("threshold_percent") or 10.0)
    version_id = body.get("version_id")

    desc_col = mapping.get("description")
    price_col = mapping.get("price")
    mat_col = mapping.get("material_price")
    lab_col = mapping.get("labor_price")
    mac_col = mapping.get("machine_price")
    dep_col = mapping.get("depreciation_pct")
    done_col = mapping.get("dep_done_pct")
    if not desc_col or not (price_col or mat_col or lab_col or mac_col):
        raise HTTPException(400, "Cần chọn cột Hạng mục/Diễn giải và cột Đơn giá (tổng hoặc VT/NC/Khấu hao)")
    write_col = price_col or mat_col or lab_col or mac_col

    rec = get_workbook_record(wb_id)
    wb = openpyxl.load_workbook(Path(rec["file_path"]), data_only=True)
    if sheet_name not in wb.sheetnames:
        raise HTTPException(400, f"Sheet {sheet_name} không tồn tại")
    sheet = wb[sheet_name]

    # Load all price library items & confirmed learning memory
    with get_connection() as conn:
        if not version_id:
            act = conn.execute("SELECT name FROM sqlite_master WHERE name='mep_library_versions'").fetchone() and \
                conn.execute("SELECT id FROM mep_library_versions WHERE status='active' ORDER BY id DESC LIMIT 1").fetchone()
            version_id = act[0] if act else None
        if version_id:
            lib_rows = conn.execute("SELECT * FROM mep_libraries WHERE version = ?", (int(version_id),)).fetchall()
        else:
            lib_rows = conn.execute("SELECT * FROM mep_libraries").fetchall()
        library_items = [dict(r) for r in lib_rows]
        mem_rows = conn.execute("SELECT * FROM mep_learning_memory").fetchall()
        learning_memory = [dict(r) for r in mem_rows]

    session_id = str(uuid.uuid4())
    review_items = []
    total_variance = 0.0
    passed_c = higher_c = lower_c = not_found_c = spec_mis_c = manual_c = 0

    max_r = sheet.max_row or 1

    for r in range(header_row + 1, max_r + 1):
        desc_val = sheet.cell(row=r, column=desc_col).value
        price_val = sheet.cell(row=r, column=price_col).value if price_col else None
        stt_val = sheet.cell(row=r, column=mapping["stt"]).value if mapping.get("stt") else ""
        boq_mat = parse_number(sheet.cell(row=r, column=mat_col).value) if mat_col else None
        boq_lab = parse_number(sheet.cell(row=r, column=lab_col).value) if lab_col else None
        boq_mac = parse_number(sheet.cell(row=r, column=mac_col).value) if mac_col else None
        boq_amt = parse_number(sheet.cell(row=r, column=mapping["amount"]).value) if mapping.get("amount") else None
        boq_dep = dep_fraction(sheet.cell(row=r, column=dep_col).value) if dep_col else None
        boq_done = dep_fraction(sheet.cell(row=r, column=done_col).value) if done_col else None
        note_val = sheet.cell(row=r, column=mapping["note"]).value if mapping.get("note") else None

        # Optional columns
        code_val = sheet.cell(row=r, column=mapping["code"]).value if mapping.get("code") else ""
        spec_val = sheet.cell(row=r, column=mapping["specification"]).value if mapping.get("specification") else ""
        brand_val = sheet.cell(row=r, column=mapping["brand"]).value if mapping.get("brand") else ""
        model_val = sheet.cell(row=r, column=mapping["model"]).value if mapping.get("model") else ""
        unit_val = sheet.cell(row=r, column=mapping["unit"]).value if mapping.get("unit") else ""
        qty_val = sheet.cell(row=r, column=mapping["quantity"]).value if mapping.get("quantity") else None

        # Parse numeric quantity & price
        boq_price = parse_number(price_val) if price_col else None
        if boq_price is None:
            # Đơn giá = VT × %KH + NC + Máy khi có cột % khấu hao vật tư
            mat_part = (boq_mat or 0.0) * boq_dep if boq_dep is not None else (boq_mat or 0.0)
            boq_price = mat_part + (boq_lab or 0.0) + (boq_mac or 0.0)
        boq_qty = parse_number(qty_val) or 0.0

        if boq_reader.is_index_row(str(desc_val or "")):
            continue

        # Filter out headers, subtotals, blank rows
        if is_non_item_row(str(desc_val or ''), str(unit_val or ''), boq_price, boq_qty):
            continue

        item_dict = {
            "code": str(code_val or "").strip(),
            "description": str(desc_val or "").strip(),
            "specification": str(spec_val or "").strip(),
            "brand": str(brand_val or "").strip(),
            "model": str(model_val or "").strip(),
            "unit": str(unit_val or "").strip(),
            "quantity": boq_qty,
            # Có %KH: so với thư viện theo giá trị đầy đủ VT + NC (+ máy), không theo giá đã khấu hao
            "price": ((boq_mat or 0.0) + (boq_lab or 0.0) + (boq_mac or 0.0)) if boq_dep is not None and boq_mat else boq_price
        }

        match_res = match_single_item(item_dict, library_items, learning_memory, threshold_percent)

        st = match_res["status"]
        if st == 'passed': passed_c += 1
        elif st == 'higher': higher_c += 1
        elif st == 'lower': lower_c += 1
        elif st == 'not_found': not_found_c += 1
        elif st == 'spec_mismatch': spec_mis_c += 1
        elif st == 'manual_check': manual_c += 1

        if boq_qty and match_res["difference"]:
            total_variance += match_res["difference"] * boq_qty

        review_items.append({
            "session_id": session_id,
            "sheet_name": sheet_name,
            "excel_row": r,
            "excel_col": write_col,
            "boq_stt": str(stt_val if stt_val is not None else "").strip(),
            "boq_material": boq_mat,
            "boq_labor": boq_lab,
            "boq_machine": boq_mac,
            "boq_dep_pct": boq_dep,
            "boq_dep_done": boq_done,
            "boq_amount": boq_amt,
            "boq_note": str(note_val).strip() if note_val not in (None, "") else None,
            "boq_code": item_dict["code"],
            "boq_description": item_dict["description"],
            "boq_spec": item_dict["specification"],
            "boq_brand": item_dict["brand"],
            "boq_model": item_dict["model"],
            "boq_unit": item_dict["unit"],
            "boq_quantity": boq_qty,
            "boq_price": boq_price,
            "lib_id": match_res["lib_id"],
            "lib_code": match_res["lib_code"],
            "lib_description": match_res["lib_description"],
            "lib_spec": match_res["lib_spec"],
            "lib_brand": match_res["lib_brand"],
            "lib_model": match_res["lib_model"],
            "lib_unit": match_res["lib_unit"],
            "lib_price": match_res["lib_price"],
            "lib_source": match_res["lib_source"],
            "difference": match_res["difference"],
            "diff_percent": match_res["diff_percent"],
            "status": st,
            "match_method": match_res["match_method"],
            "match_score": match_res["match_score"],
            "warning_color": match_res["warning_color"],
            "user_action": "none",
            "user_note": match_res.get("user_note", ""),
            "candidates_json": json.dumps(match_res.get("candidates", []), ensure_ascii=False),
            "updated_at": now_iso()
        })

    wb.close()

    # Save to SQLite
    with get_connection() as conn:
        conn.execute("""
            INSERT INTO mep_review_sessions (
                id, workbook_id, sheet_name, threshold_percent, mapping_config,
                total_items, passed_count, higher_count, lower_count, not_found_count,
                spec_mismatch_count, manual_check_count, total_variance, status, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            session_id, wb_id, sheet_name, threshold_percent, json.dumps(mapping, ensure_ascii=False),
            len(review_items), passed_c, higher_c, lower_c, not_found_c,
            spec_mis_c, manual_c, round(total_variance, 2), "completed", now_iso()
        ))

        for it in review_items:
            conn.execute("""
                INSERT INTO mep_review_items (
                    session_id, sheet_name, excel_row, excel_col, boq_code, boq_description,
                    boq_spec, boq_brand, boq_model, boq_unit, boq_quantity, boq_price,
                    lib_id, lib_code, lib_description, lib_spec, lib_brand, lib_model,
                    lib_unit, lib_price, lib_source, difference, diff_percent, status,
                    match_method, match_score, warning_color, user_action, user_note,
                    candidates_json, updated_at, boq_stt, boq_material, boq_labor,
                    boq_machine, boq_amount, boq_note, boq_dep_pct, boq_dep_done
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                it["session_id"], it["sheet_name"], it["excel_row"], it["excel_col"], it["boq_code"], it["boq_description"],
                it["boq_spec"], it["boq_brand"], it["boq_model"], it["boq_unit"], it["boq_quantity"], it["boq_price"],
                it["lib_id"], it["lib_code"], it["lib_description"], it["lib_spec"], it["lib_brand"], it["lib_model"],
                it["lib_unit"], it["lib_price"], it["lib_source"], it["difference"], it["diff_percent"], it["status"],
                it["match_method"], it["match_score"], it["warning_color"], it["user_action"], it["user_note"],
                it["candidates_json"], it["updated_at"], it["boq_stt"], it["boq_material"], it["boq_labor"],
                it["boq_machine"], it["boq_amount"], it["boq_note"], it["boq_dep_pct"], it.get("boq_dep_done")
            ))

    return {
        "session_id": session_id,
        "workbook_id": wb_id,
        "sheet_name": sheet_name,
        "threshold_percent": threshold_percent,
        "summary": {
            "total": len(review_items),
            "passed": passed_c,
            "higher": higher_c,
            "lower": lower_c,
            "not_found": not_found_c,
            "spec_mismatch": spec_mis_c,
            "manual_check": manual_c,
            "total_variance": round(total_variance, 2)
        }
    }

@router.get("/review/session/{session_id}")
def get_review_session_details(session_id: str):
    """Retrieves session summary and review item rows."""
    with get_connection() as conn:
        session = conn.execute("SELECT * FROM mep_review_sessions WHERE id = ?", (session_id,)).fetchone()
        if not session:
            raise HTTPException(404, "Không tìm thấy phiên soát xét")
        items = conn.execute("SELECT * FROM mep_review_items WHERE session_id = ? ORDER BY excel_row", (session_id,)).fetchall()
        
        parsed_items = []
        for r in items:
            d = dict(r)
            if d.get("candidates_json"):
                try: d["candidates"] = json.loads(d["candidates_json"])
                except: d["candidates"] = []
            parsed_items.append(d)

        return {
            "session": dict(session),
            "items": parsed_items
        }

@router.post("/review/item/{item_id}/action")
def update_review_item_action(item_id: int, body: Dict[str, Any]):
    """Applies user manual review decision: apply library price, keep BOQ price, or edit custom price."""
    action = body.get("action") # 'applied_lib', 'kept_boq', 'edited_custom', 'marked_checked'
    user_note = body.get("user_note", "")
    edited_price = body.get("edited_price")
    remember_memory = bool(body.get("remember_memory"))

    with get_connection() as conn:
        item = conn.execute("SELECT * FROM mep_review_items WHERE id = ?", (item_id,)).fetchone()
        if not item:
            raise HTTPException(404, "Không tìm thấy dòng soát xét")

        final_price = item["boq_price"]
        if action == 'applied_lib' and item["lib_price"]:
            final_price = item["lib_price"]
        elif action == 'edited_custom' and edited_price is not None:
            final_price = float(edited_price)

        # Update Excel cell directly if price changed
        if final_price != item["boq_price"]:
            sess = conn.execute("SELECT workbook_id FROM mep_review_sessions WHERE id = ?", (item["session_id"],)).fetchone()
            if sess:
                update_sheet_cells(sess["workbook_id"], item["sheet_name"], [{
                    "r": item["excel_row"],
                    "c": item["excel_col"],
                    "v": final_price
                }])

        # Update item record
        new_color = 'green' if action in ['applied_lib', 'marked_checked', 'kept_boq'] else item["warning_color"]
        conn.execute("""
            UPDATE mep_review_items SET
                user_action = ?, user_note = ?, edited_price = ?,
                warning_color = ?, updated_at = ?
            WHERE id = ?
        """, (action, user_note, final_price, new_color, now_iso(), item_id))

        # Add to learning memory if requested
        if remember_memory and item["lib_description"]:
            conn.execute("""
                INSERT INTO mep_learning_memory (
                    boq_code, boq_description, boq_unit, library_code, library_description,
                    library_unit, confirmed_price, notes, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                item["boq_code"], item["boq_description"], item["boq_unit"],
                item["lib_code"], item["lib_description"], item["lib_unit"],
                final_price, user_note or "Người dùng xác nhận tương đương", now_iso()
            ))

    return {"updated": True, "final_price": final_price}

# ==================== 4. LEARNING MEMORY ====================

@router.get("/review/learning-memory")
def list_learning_memory():
    """Lists all confirmed material alias pairs."""
    with get_connection() as conn:
        rows = conn.execute("SELECT * FROM mep_learning_memory ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]

@router.delete("/review/learning-memory/{mem_id}")
def delete_learning_memory(mem_id: int):
    """Revokes a learned alias pair."""
    with get_connection() as conn:
        conn.execute("DELETE FROM mep_learning_memory WHERE id = ?", (mem_id,))
    return {"deleted": True, "id": mem_id}

# ==================== 5. REPORTS & HISTORY ====================

@router.get("/reports/variance/{session_id}")
def download_variance_report(session_id: str):
    """Exports 16-column Excel variance report."""
    try:
        report_path = generate_variance_report(session_id)
        return FileResponse(report_path, filename=f"Bao-cao-chenh-lech-don-gia-{session_id[:8]}.xlsx")
    except Exception as e:
        raise HTTPException(500, f"Lỗi tạo báo cáo chênh lệch: {str(e)}")

@router.get("/history/sessions")
def list_review_sessions():
    """Lists all review sessions."""
    with get_connection() as conn:
        rows = conn.execute("""
            SELECT s.*, w.original_filename
            FROM mep_review_sessions s
            LEFT JOIN mep_workbooks w ON s.workbook_id = w.id
            ORDER BY s.created_at DESC
        """).fetchall()
        return [dict(r) for r in rows]

@router.get("/overview/stats")
def get_overview_statistics():
    """Returns dashboard KPIs for Overview tab."""
    with get_connection() as conn:
        lib_count = conn.execute("SELECT COUNT(*) FROM mep_libraries").fetchone()[0]
        sys_counts = conn.execute("SELECT system, COUNT(*) as cnt FROM mep_libraries GROUP BY system").fetchall()
        wb_count = conn.execute("SELECT COUNT(*) FROM mep_workbooks").fetchone()[0]
        sess_count = conn.execute("SELECT COUNT(*) FROM mep_review_sessions").fetchone()[0]
        mem_count = conn.execute("SELECT COUNT(*) FROM mep_learning_memory").fetchone()[0]
        latest_sess = conn.execute("SELECT * FROM mep_review_sessions ORDER BY created_at DESC LIMIT 1").fetchone()

        return {
            "library_total": lib_count,
            "library_by_system": {r["system"]: r["cnt"] for r in sys_counts},
            "workbooks_total": wb_count,
            "sessions_total": sess_count,
            "memory_total": mem_count,
            "latest_session": dict(latest_sess) if latest_sess else None
        }

