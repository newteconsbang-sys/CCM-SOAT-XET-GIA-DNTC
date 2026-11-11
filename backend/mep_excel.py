"""Excel Workspace Backend: openpyxl processing, multi-sheet rendering,
safe cell updating, styling, and variance report generation.
"""
import uuid
import json
import os
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from backend.mep_db import get_connection, now_iso

ROOT = Path(__file__).resolve().parents[1]
UPLOADS_DIR = ROOT / "uploads"
BACKUPS_DIR = ROOT / "backups"
REPORTS_DIR = ROOT / "reports"

for folder in [UPLOADS_DIR, BACKUPS_DIR, REPORTS_DIR]:
    folder.mkdir(exist_ok=True)

# Color Fills for Highlights
FILL_RED = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
FONT_RED = Font(color="991B1B", bold=True)

FILL_YELLOW = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
FONT_YELLOW = Font(color="92400E", bold=True)

FILL_GREEN = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
FONT_GREEN = Font(color="166534", bold=True)

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)

def save_uploaded_workbook(file_bytes: bytes, original_filename: str) -> Dict[str, Any]:
    """Saves uploaded workbook, stores backup, inspects sheets and macro info."""
    ext = Path(original_filename).suffix.lower()
    if ext not in ['.xlsx', '.xlsm']:
        raise ValueError("Chỉ hỗ trợ file Excel định dạng .xlsx hoặc .xlsm")

    wb_id = str(uuid.uuid4())
    stored_name = f"{wb_id}{ext}"
    file_path = UPLOADS_DIR / stored_name
    backup_path = BACKUPS_DIR / f"{wb_id}_original{ext}"

    file_path.write_bytes(file_bytes)
    shutil.copy2(file_path, backup_path)

    # Inspect with openpyxl
    keep_vba = (ext == '.xlsm')
    wb = openpyxl.load_workbook(file_path, read_only=False, keep_vba=keep_vba, data_only=False)
    
    sheets_info = []
    for name in wb.sheetnames:
        sheet = wb[name]
        sheets_info.append({
            "name": name,
            "max_row": sheet.max_row or 1,
            "max_col": sheet.max_column or 1
        })
    wb.close()

    sheets_json = json.dumps(sheets_info, ensure_ascii=False)
    with get_connection() as conn:
        conn.execute("""
            INSERT INTO mep_workbooks (id, filename, original_filename, file_path, backup_path, file_type, sheets_info, uploaded_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (wb_id, stored_name, original_filename, str(file_path), str(backup_path), ext, sheets_json, now_iso()))

    return {
        "id": wb_id,
        "filename": original_filename,
        "file_type": ext,
        "is_macro": ext == '.xlsm',
        "sheets": sheets_info,
        "uploaded_at": now_iso()
    }

def get_workbook_record(wb_id: str) -> Dict[str, Any]:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM mep_workbooks WHERE id = ?", (wb_id,)).fetchone()
        if not row:
            raise ValueError("Không tìm thấy workbook")
        return dict(row)

def get_sheet_grid(wb_id: str, sheet_name: Optional[str] = None, max_rows: int = 500, max_cols: int = 35) -> Dict[str, Any]:
    """Extracts cells, values, formatting, merges, and warning highlights."""
    rec = get_workbook_record(wb_id)
    file_path = Path(rec["file_path"])
    keep_vba = (rec["file_type"] == '.xlsm')

    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)
    sheet_names = wb.sheetnames
    target_sheet = sheet_name if sheet_name and sheet_name in sheet_names else sheet_names[0]
    sheet = wb[target_sheet]

    # Check for active review items to show highlights
    review_highlights = {}
    with get_connection() as conn:
        rows = conn.execute("""
            SELECT excel_row, excel_col, warning_color, status, boq_price, lib_price, difference, diff_percent, lib_source, user_action, user_note, id
            FROM mep_review_items
            WHERE sheet_name = ? AND session_id IN (
                SELECT id FROM mep_review_sessions WHERE workbook_id = ? ORDER BY created_at DESC LIMIT 1
            )
        """, (target_sheet, wb_id)).fetchall()
        for r in rows:
            key = f"{r['excel_row']}_{r['excel_col']}"
            review_highlights[key] = dict(r)

    row_count = min(sheet.max_row or 1, max_rows)
    col_count = min(sheet.max_column or 1, max_cols)

    cells = {}
    for r in range(1, row_count + 1):
        for c in range(1, col_count + 1):
            cell = sheet.cell(row=r, column=c)
            val = cell.value
            formula = None
            if isinstance(val, str) and val.startswith('='):
                formula = val

            # Style attributes
            bg = None
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                rgb = str(cell.fill.start_color.rgb)
                if rgb and rgb not in ['00000000', 'None']:
                    bg = f"#{rgb[-6:]}"

            fg = None
            bold = False
            italic = False
            if cell.font:
                bold = bool(cell.font.bold)
                italic = bool(cell.font.italic)
                if cell.font.color and cell.font.color.rgb:
                    c_rgb = str(cell.font.color.rgb)
                    if c_rgb and c_rgb not in ['00000000', 'None']:
                        fg = f"#{c_rgb[-6:]}"

            align = 'left'
            if cell.alignment and cell.alignment.horizontal:
                align = cell.alignment.horizontal

            cell_key = f"{r}_{c}"
            highlight = review_highlights.get(cell_key)

            cells[cell_key] = {
                "r": r,
                "c": c,
                "v": val,
                "f": formula,
                "bg": bg,
                "fg": fg,
                "b": bold,
                "i": italic,
                "align": align,
                "warn": highlight.get("warning_color") if highlight else None,
                "review": highlight
            }

    # Merged ranges
    merges = []
    for rng in sheet.merged_cells.ranges:
        if rng.min_row <= row_count and rng.min_col <= col_count:
            merges.append({
                "s_r": rng.min_row,
                "s_c": rng.min_col,
                "e_r": min(rng.max_row, row_count),
                "e_c": min(rng.max_col, col_count)
            })

    wb.close()

    return {
        "workbook_id": wb_id,
        "active_sheet": target_sheet,
        "sheet_names": sheet_names,
        "row_count": row_count,
        "col_count": col_count,
        "total_rows": sheet.max_row or 1,
        "total_cols": sheet.max_column or 1,
        "cells": cells,
        "merges": merges
    }

def update_sheet_cells(wb_id: str, sheet_name: str, updates: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Applies cell modifications from Web Editor into Excel workbook."""
    rec = get_workbook_record(wb_id)
    file_path = Path(rec["file_path"])
    keep_vba = (rec["file_type"] == '.xlsm')

    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet {sheet_name} không tồn tại")

    sheet = wb[sheet_name]
    for u in updates:
        r = int(u["r"])
        c = int(u["c"])
        val = u.get("v")
        # Try numeric conversion if possible
        if isinstance(val, str) and not val.startswith('='):
            clean_str = val.replace(',', '').strip()
            try:
                if '.' in clean_str:
                    val = float(clean_str)
                else:
                    val = int(clean_str)
            except ValueError:
                pass
        sheet.cell(row=r, column=c, value=val)

    wb.save(file_path)
    wb.close()
    return {"saved": True, "updated_cells": len(updates)}

def modify_sheet_structure(wb_id: str, sheet_name: str, action: str, index: int = 1, extra_name: Optional[str] = None) -> Dict[str, Any]:
    """Inserts or deletes rows and columns, or adds/renames sheets."""
    rec = get_workbook_record(wb_id)
    file_path = Path(rec["file_path"])
    keep_vba = (rec["file_type"] == '.xlsm')

    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)
    
    if action == 'add_sheet':
        new_name = extra_name or f"Sheet{len(wb.sheetnames)+1}"
        wb.create_sheet(title=new_name)
    elif action == 'delete_sheet':
        if len(wb.sheetnames) <= 1:
            raise ValueError("Không thể xóa sheet duy nhất trong workbook")
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
    elif action == 'rename_sheet' and extra_name:
        if sheet_name in wb.sheetnames:
            wb[sheet_name].title = extra_name
    elif sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        if action == 'insert_row':
            sheet.insert_rows(index, 1)
        elif action == 'delete_row':
            sheet.delete_rows(index, 1)
        elif action == 'insert_col':
            sheet.insert_cols(index, 1)
        elif action == 'delete_col':
            sheet.delete_cols(index, 1)

    wb.save(file_path)
    sheet_names = wb.sheetnames
    wb.close()
    return {"success": True, "action": action, "sheet_names": sheet_names}

def apply_review_highlights(wb_id: str, session_id: str) -> str:
    """Highlights flagged cells directly in workbook and attaches comments."""
    rec = get_workbook_record(wb_id)
    file_path = Path(rec["file_path"])
    keep_vba = (rec["file_type"] == '.xlsm')

    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)

    with get_connection() as conn:
        items = conn.execute("""
            SELECT * FROM mep_review_items WHERE session_id = ?
        """, (session_id,)).fetchall()

    for item in items:
        s_name = item["sheet_name"]
        if s_name not in wb.sheetnames:
            continue
        ws = wb[s_name]
        r = item["excel_row"]
        c = item["excel_col"]
        cell = ws.cell(row=r, column=c)

        status = item["status"]
        if status == 'higher':
            cell.fill = FILL_RED
            cell.font = FONT_RED
        elif status in ['lower', 'manual_check', 'spec_mismatch', 'not_found']:
            cell.fill = FILL_YELLOW
            cell.font = FONT_YELLOW
        elif status == 'passed':
            cell.fill = FILL_GREEN
            cell.font = FONT_GREEN

        # Add Excel Comment
        boq_p = f"{item['boq_price']:,.0f}" if item['boq_price'] else "0"
        lib_p = f"{item['lib_price']:,.0f}" if item['lib_price'] else "Chưa có"
        diff_p = f"{item['difference']:,.0f}" if item['difference'] else "0"
        diff_pct = f"{item['diff_percent']:+.1f}%" if item['diff_percent'] is not None else "--"
        
        status_vi = {
            'passed': 'ĐẠT (Trong ngưỡng cho phép)',
            'higher': 'CẢNH BÁO: GIÁ CAO HƠN THƯ VIỆN',
            'lower': 'LƯU Ý: GIÁ THẤP HƠN THƯ VIỆN',
            'not_found': 'CHƯA TÌM THẤY GIÁ THAM CHIẾU',
            'spec_mismatch': 'LƯU Ý: KHÔNG KHỚP THÔNG SỐ',
            'manual_check': 'CẦN XÁC NHẬN THỦ CÔNG'
        }.get(status, status)

        comment_text = (
            f"[CCM-AI SOÁT XÉT ĐƠN GIÁ MEP]\n"
            f"Trạng thái: {status_vi}\n"
            f"Giá BOQ: {boq_p} đ\n"
            f"Giá thư viện: {lib_p} đ\n"
            f"Chênh lệch: {diff_p} đ ({diff_pct})\n"
            f"Nguồn: {item['lib_source'] or 'N/A'}"
        )
        cell.comment = Comment(comment_text, "CCM-AI Price Review")

    reviewed_filename = f"{wb_id}_reviewed{rec['file_type']}"
    reviewed_path = REPORTS_DIR / reviewed_filename
    wb.save(reviewed_path)
    wb.close()
    return str(reviewed_path)

def generate_variance_report(session_id: str) -> str:
    """Generates the 16-column MEP Price Variance Report in Excel."""
    with get_connection() as conn:
        session = conn.execute("SELECT * FROM mep_review_sessions WHERE id = ?", (session_id,)).fetchone()
        if not session:
            raise ValueError("Không tìm thấy phiên soát xét")
        items = conn.execute("SELECT * FROM mep_review_items WHERE session_id = ? ORDER BY excel_row", (session_id,)).fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Báo cáo chênh lệch đơn giá"
    ws.views.sheetView[0].showGridLines = True

    # Title block
    ws.merge_cells("A1:P1")
    title_cell = ws["A1"]
    title_cell.value = "BÁO CÁO SOÁT XÉT ĐƠN GIÁ MEP – CCM-AI PRICE REVIEW"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # Metadata block
    ws["A2"] = f"Phiên: {session['id'][:8]}"
    ws["C2"] = f"Ngày soát xét: {session['created_at'][:19].replace('T', ' ')}"
    ws["G2"] = f"Ngưỡng cảnh báo: ±{session['threshold_percent']}%"
    ws["K2"] = f"Tổng số đầu mục: {session['total_items']}"
    for cell_pos in ["A2", "C2", "G2", "K2"]:
        ws[cell_pos].font = Font(name="Calibri", size=10, italic=True)

    # KPI Summary Row
    ws["A3"] = "Tổng hợp kết quả:"
    ws["C3"] = f"🟢 Đạt: {session['passed_count']}"
    ws["E3"] = f"🔴 Cao hơn: {session['higher_count']}"
    ws["G3"] = f"🟡 Thấp hơn: {session['lower_count']}"
    ws["I3"] = f"⚪ Chưa có giá: {session['not_found_count']}"
    ws["K3"] = f"⚠️ Lệch thông số: {session['spec_mismatch_count']}"
    ws["M3"] = f"Tổng chênh lệch: {session['total_variance']:,.0f} đ"
    for col_letter in ["A", "C", "E", "G", "I", "K", "M"]:
        ws[f"{col_letter}3"].font = Font(name="Calibri", size=11, bold=True)
    ws.row_dimensions[3].height = 25

    headers = [
        "STT", "Sheet", "Dòng Excel", "Mã vật tư", "Mô tả BOQ", "Thông số kỹ thuật",
        "ĐVT", "Khối lượng", "Đơn giá BOQ (đ)", "Đơn giá Thư viện (đ)", "Chênh lệch (đ)",
        "Tỷ lệ (%)", "Nguồn tham chiếu", "Kết quả kiểm tra", "Trạng thái xử lý", "Ghi chú"
    ]

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[5].height = 30
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=5, column=col_idx, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center_align
        c.border = THIN_BORDER

    # Status mapping labels
    status_map = {
        'passed': ('Đạt', FILL_GREEN, FONT_GREEN),
        'higher': ('Giá cao hơn thư viện', FILL_RED, FONT_RED),
        'lower': ('Giá thấp hơn thư viện', FILL_YELLOW, FONT_YELLOW),
        'not_found': ('Chưa tìm thấy giá', PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"), Font(color="475569")),
        'spec_mismatch': ('Không khớp thông số', FILL_YELLOW, FONT_YELLOW),
        'manual_check': ('Cần xác nhận thủ công', FILL_YELLOW, FONT_YELLOW)
    }

    row_num = 6
    for idx, it in enumerate(items, start=1):
        ws.row_dimensions[row_num].height = 22
        st_label, st_fill, st_font = status_map.get(it["status"], (it["status"], None, None))

        row_data = [
            idx,
            it["sheet_name"],
            it["excel_row"],
            it["boq_code"] or "",
            it["boq_description"],
            it["boq_spec"] or "",
            it["boq_unit"] or "",
            it["boq_quantity"],
            it["boq_price"],
            it["lib_price"],
            it["difference"],
            (it["diff_percent"] / 100.0) if it["diff_percent"] is not None else None,
            it["lib_source"] or "",
            st_label,
            it["user_action"] or "Chưa xử lý",
            it["user_note"] or ""
        ]

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.border = THIN_BORDER

            # Formatting numbers
            if col_idx in [1, 3]:
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in [8]:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx in [9, 10, 11]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 12:
                cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 14 and st_fill:
                cell.fill = st_fill
                cell.font = st_font
                cell.alignment = Alignment(horizontal="center")

        row_num += 1

    # Auto column widths
    col_widths = {1: 6, 2: 14, 3: 10, 4: 14, 5: 35, 6: 22, 7: 8, 8: 12, 9: 16, 10: 16, 11: 16, 12: 12, 13: 25, 14: 20, 15: 16, 16: 25}
    for c_idx, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_idx)].width = w

    report_filename = f"Bao-cao-chenh-lech-don-gia-{session_id[:8]}.xlsx"
    report_path = REPORTS_DIR / report_filename
    wb.save(report_path)
    wb.close()
    return str(report_path)

