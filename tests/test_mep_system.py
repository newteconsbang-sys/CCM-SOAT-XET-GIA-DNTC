"""Comprehensive End-to-End Test for CCM-AI MEP Price Review System."""
import sys
import json
from pathlib import Path

# Force UTF-8 stdout
sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'backend' / 'vendor'))
sys.path.insert(0, str(ROOT))

import openpyxl
from backend.mep_db import get_connection, now_iso
from backend.mep_excel import save_uploaded_workbook, get_sheet_grid, update_sheet_cells, apply_review_highlights, generate_variance_report
from backend.mep_engine import match_single_item, normalize_mep_description, extract_mep_specs
from backend.app import app
from fastapi.testclient import TestClient

client = TestClient(app, base_url="https://localhost:3443")

def test_full_mep_workflow():
    print("=== Step 1: Health Check ===")
    res = client.get("/api/health")
    assert res.status_code == 200, f"Health check failed: {res.text}"
    print("Health response:", res.json())

    # Get session token
    session_res = client.get("/api/session")
    token = session_res.json()["token"]
    headers = {"X-CCM-Token": token}

    print("\n=== Step 2: Check Price Library Data ===")
    lib_res = client.get("/api/mep/libraries", headers=headers)
    assert lib_res.status_code == 200
    lib_data = lib_res.json()
    print(f"Total library items: {lib_data['total']}")
    assert lib_data['total'] > 0, "Price library should contain seeded items"

    print("\n=== Step 3: Upload Sample BOQ ===")
    sample_boq_path = ROOT / "BoQ-MEP-Mau-Can-Soat-Gia.xlsx"
    assert sample_boq_path.exists(), "Sample BOQ file not found"
    
    with open(sample_boq_path, "rb") as f:
        upload_res = client.post(
            "/api/mep/workbooks/upload",
            files={"file": ("BoQ-MEP-Mau-Can-Soat-Gia.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            headers=headers
        )
    assert upload_res.status_code == 200, f"Upload failed: {upload_res.text}"
    wb_info = upload_res.json()
    wb_id = wb_info["id"]
    print("Uploaded workbook ID:", wb_id, "Sheets:", [s["name"] for s in wb_info["sheets"]])

    print("\n=== Step 4: Inspect Excel Workspace Grid ===")
    grid_res = client.get(f"/api/mep/workbooks/{wb_id}/sheet?sheet_name=BoQ Điện Nước MEP", headers=headers)
    assert grid_res.status_code == 200
    grid = grid_res.json()
    print(f"Grid loaded: {grid['row_count']} rows x {grid['col_count']} cols, {len(grid['cells'])} cells populated")

    print("\n=== Step 5: Auto-detect Column Mapping ===")
    map_res = client.post(
        "/api/mep/review/suggest-mapping",
        json={"workbook_id": wb_id, "sheet_name": "BoQ Điện Nước MEP", "header_row": 2},
        headers=headers
    )
    assert map_res.status_code == 200
    suggested = map_res.json()
    print("Auto-detected mapping:", suggested["mapping"])
    mapping = suggested["mapping"]
    assert mapping["description"] is not None, "Description column must be mapped"
    assert mapping["price"] is not None, "Price column must be mapped"

    print("\n=== Step 6: Run Price Review Engine ===")
    review_res = client.post(
        "/api/mep/review/run",
        json={
            "workbook_id": wb_id,
            "sheet_name": "BoQ Điện Nước MEP",
            "header_row": 2,
            "mapping": mapping,
            "threshold_percent": 10.0
        },
        headers=headers
    )
    assert review_res.status_code == 200, f"Review failed: {review_res.text}"
    review_data = review_res.json()
    sess_id = review_data["session_id"]
    summary = review_data["summary"]
    print("Review Summary:", summary)
    assert summary["total"] > 0, "Review items count should be > 0"
    print(f"🟢 Passed: {summary['passed']}, 🔴 Higher: {summary['higher']}, 🟡 Lower: {summary['lower']}, ⚪ Not found: {summary['not_found']}")

    print("\n=== Step 7: Get Review Item Details ===")
    detail_res = client.get(f"/api/mep/review/session/{sess_id}", headers=headers)
    assert detail_res.status_code == 200
    items = detail_res.json()["items"]
    print(f"Retrieved {len(items)} review items")
    for it in items[:5]:
        print(f"  - Dòng {it['excel_row']}: {it['boq_description']} | BOQ: {it['boq_price']:,} đ | Lib: {it['lib_price']} | Status: {it['status']} | Diff: {it['diff_percent']}% | Color: {it['warning_color']}")

    print("\n=== Step 8: Apply User Action (Apply Reference Price) ===")
    higher_item = next((it for it in items if it["status"] == "higher" and it["lib_price"]), None)
    if higher_item:
        print(f"Applying library price for item id {higher_item['id']} ({higher_item['boq_description']})...")
        action_res = client.post(
            f"/api/mep/review/item/{higher_item['id']}/action",
            json={
                "action": "applied_lib",
                "user_note": "Người dùng chấp thuận giá thư viện",
                "remember_memory": True
            },
            headers=headers
        )
        assert action_res.status_code == 200
        print("Action applied successfully. New cell price:", action_res.json()["final_price"])

    print("\n=== Step 9: Verify Learning Memory Was Updated ===")
    mem_res = client.get("/api/mep/review/learning-memory", headers=headers)
    assert mem_res.status_code == 200
    memories = mem_res.json()
    print(f"Total learning memories: {len(memories)}")
    assert len(memories) > 0, "Memory should contain user-confirmed alias"

    print("\n=== Step 10: Generate and Export Reviewed BOQ File ===")
    reviewed_wb_path = apply_review_highlights(wb_id, sess_id)
    assert Path(reviewed_wb_path).exists(), "Reviewed workbook file must exist"
    wb_reviewed = openpyxl.load_workbook(reviewed_wb_path, data_only=False)
    assert "BoQ Điện Nước MEP" in wb_reviewed.sheetnames
    ws_rev = wb_reviewed["BoQ Điện Nước MEP"]
    # Check that cell has comment
    has_comment = any(cell.comment is not None for row in ws_rev.iter_rows() for cell in row)
    assert has_comment, "Reviewed workbook should contain inspection comments"
    wb_reviewed.close()
    print("Reviewed BOQ generated successfully with highlights and comments:", reviewed_wb_path)

    print("\n=== Step 11: Generate and Export 16-Column Variance Report ===")
    variance_rep_path = generate_variance_report(sess_id)
    assert Path(variance_rep_path).exists(), "Variance report file must exist"
    wb_rep = openpyxl.load_workbook(variance_rep_path)
    ws_rep = wb_rep.active
    assert ws_rep["A1"].value is not None, "Report title should exist"
    assert ws_rep["A5"].value == "STT", "Column 1 must be STT"
    assert ws_rep["P5"].value == "Ghi chú", "Column 16 must be Ghi chú"
    print(f"Variance report valid: {ws_rep.max_row} rows x {ws_rep.max_column} columns")
    wb_rep.close()

    print("\n=== Step 12: Overview Dashboard Stats ===")
    stats_res = client.get("/api/mep/overview/stats", headers=headers)
    assert stats_res.status_code == 200
    stats = stats_res.json()
    print("Overview Stats:", stats)
    assert stats["workbooks_total"] > 0
    assert stats["sessions_total"] > 0

    print("\n ALL 12 END-TO-END TESTS PASSED PERFECTLY! ")

if __name__ == "__main__":
    test_full_mep_workflow()
