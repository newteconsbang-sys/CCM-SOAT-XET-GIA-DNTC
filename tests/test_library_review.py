"""Kiểm thử Thư viện giá (version, sửa, so sánh) và Soát xét giá (lưới 2 phần, lịch sử, xuất file gốc).

Chạy trên database TẠM – không đụng data/ccm.sqlite3 thật:
    runtime\\python\\python.exe -m pytest -q tests\\test_library_review.py
"""
import io
import json
import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from backend import mep_db  # noqa: E402
from fastapi import FastAPI  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402
from backend import mep_router as mep_router_mod, mep_packages, library_service, review_service, library_structure, review_sheets  # noqa: E402
from backend.mep_router import router as mep_router  # noqa: E402
from backend.mep_packages import router as packages_router  # noqa: E402
from backend.library_router import router as library_router  # noqa: E402
from backend.xlsx_utils import parse_number  # noqa: E402

_REAL_DB = mep_db.DB_PATH


def setup_module(_):
    """Chuyển sang database tạm cho cả module, tạo đủ bảng."""
    mep_db.DB_PATH = Path(tempfile.mkdtemp(prefix="ccm-test-")) / "test.sqlite3"
    mep_db.init_mep_db()
    mep_router_mod._ensure_review_columns()
    mep_packages.init_package_tables()
    library_service.ensure_library_schema()
    review_service.ensure_schema()
    library_structure.ensure_schema()
    review_sheets.ensure_schema()
    from backend import review_versions
    review_versions.ensure_schema()


def teardown_module(_):
    mep_db.DB_PATH = _REAL_DB

app = FastAPI()
for r in (mep_router, packages_router, library_router):
    app.include_router(r)
client = TestClient(app, headers={"X-CCM-User": "Test"})


def _xlsx(sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


LIB1 = _xlsx({
    "Vat tu": [["Mã", "Tên vật tư", "ĐVT", "Đơn giá VT", "Đơn giá NC", "Đơn giá tổng"],
               ["CD-01", "Cáp điện CV 3x4mm2", "m", "18.500", "6.500", None],
               ["ON-01", "Ống PVC D25 luồn dây", "m", 9800, 4500, 14300],
               ["X-01", "Thiếu ĐVT", "", 1, 1, None]],
    "Nhan cong": [["Mã", "Tên công việc", "ĐVT", "Đơn giá"], ["NC-01", "Nhân công điện", "công", 450000]],
    "Ca may": [["Mã", "Tên máy", "ĐVT", "Đơn giá"], ["CM-01", "Máy khoan", "ca", 285000]],
})
LIB2 = _xlsx({"Vat tu": [["Mã", "Tên vật tư", "ĐVT", "Đơn giá VT", "Đơn giá NC"],
                         ["CD-01", "Cáp điện CV 3x4mm2", "m", 19500, 6500],
                         ["TB-01", "Đèn LED panel", "bộ", 500000, 85000]]})


def _ntp():
    wb = openpyxl.Workbook()
    wb.active.title = "Bia"
    wb["Bia"]["A2"] = "=1+1"
    ws = wb.create_sheet("BG")
    ws["A1"] = "BÁO GIÁ"
    ws.merge_cells("A1:H1")
    for i, h in enumerate(["STT", "Hạng mục", "ĐVT", "Khối lượng", "Đơn giá", None, None, "Thành tiền"], 1):
        ws.cell(row=3, column=i, value=h)
    ws.merge_cells("E3:G3")
    for c, h in zip("EFG", ["Vật tư", "Nhân công", "Khấu hao"]):
        ws[f"{c}4"] = h
    rows = [("1", "Cáp điện CV 3x4mm2", "m", 100, 22000, 7000, None), ("2", "Ống PVC D25 luồn dây", "m", 50, 9800, 4500, 200)]
    for r, (stt, d, u, q, vt, nc, kh) in enumerate(rows, 5):
        ws.append([stt, d, u, q, vt, nc, kh, f"=D{r}*SUM(E{r}:G{r})"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_parse_number_vietnamese():
    assert parse_number("1.250.000") == 1250000
    assert parse_number("25.000,5") == 25000.5
    assert parse_number("18.500 đ") == 18500
    assert parse_number("abc") is None


def test_library_versions_edit_compare_export():
    r = client.post("/api/mep/library/versions/import", files={"file": ("lib1.xlsx", LIB1)})
    assert r.status_code == 200, r.text
    j = r.json()
    assert j["valid_count"] == 4 and len(j["errors"]) == 1
    assert j["by_type"] == {"Vật tư": 2, "Nhân công": 1, "Ca máy": 1}
    v1 = j["version"]["id"]
    items = client.get(f"/api/mep/library/versions/{v1}/items").json()
    cd = next(i for i in items["items"] if i["code"] == "CD-01")
    assert cd["total_price"] == 25000  # tự cộng VT + NC
    # sửa trực tiếp → tự tính lại tổng, ghi lịch sử
    up = client.put(f"/api/mep/library/items/{cd['id']}", json={"labor_price": "7.000"}).json()
    assert up["item"]["total_price"] == 25500
    assert client.put(f"/api/mep/library/items/{cd['id']}", json={"labor_price": "x"}).status_code == 400
    hist = client.get(f"/api/mep/library/versions/{v1}/history").json()
    assert any(h["field"] == "labor_price" and h["changed_by"] == "Test" for h in hist)
    # version 2 → so sánh
    v2 = client.post("/api/mep/library/versions/import", files={"file": ("lib2.xlsx", LIB2)}).json()["version"]["id"]
    cmp = client.get(f"/api/mep/library/compare?base={v1}&target={v2}").json()
    assert cmp["summary"]["added"] == 1 and cmp["summary"]["changed"] == 1 and cmp["summary"]["removed"] == 3
    versions = client.get("/api/mep/library/versions").json()
    assert [v["status"] for v in versions if v["id"] == v2] == ["active"]
    # Version nạp từ file: xuất lại đúng file gốc (đủ 3 sheet), ô đã sửa được ghi đè
    wb = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/library/versions/{v1}/export").content))
    assert wb.sheetnames == ["Vat tu", "Nhan cong", "Ca may"]
    assert wb["Vat tu"]["E2"].value == 7000
    # Có bộ lọc thì xuất bảng chuẩn theo bộ lọc
    wb = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/library/versions/{v1}/export?type=Vật tư").content))
    assert wb["Thu vien gia"].max_row == 3
    assert client.get(f"/api/mep/library/compare/export?base={v1}&target={v2}").status_code == 200
    # quay lại v1 làm hiện hành cho test soát xét
    client.post(f"/api/mep/library/versions/{v1}/activate")


def test_review_edit_history_and_export_original():
    pr = client.post("/api/mep/projects", json={"name": "Dự án test"}).json()
    r = client.post("/api/mep/packages", data={"project_id": pr["id"], "code": "T-01", "name": "Gói test", "start": "true"},
                    files={"file": ("ntp.xlsx", _ntp())}).json()
    assert r["warning"] is None
    pid = r["id"]
    d = client.get(f"/api/mep/packages/{pid}").json()
    assert d["sheet_name"] == "BG" and d["header_row"] == 4
    cap = next(i for i in d["items"] if i["boq_description"].startswith("Cáp"))
    assert cap["boq_machine"] is None and cap["boq_price"] == 29000
    assert cap["status"] == "higher" and cap["rv_material"] == 18500 and cap["review_status"] == "Cao"
    ong = next(i for i in d["items"] if i["boq_description"].startswith("Ống"))
    assert ong["boq_machine"] == 200 and ong["boq_price"] == 14500
    # sửa trực tiếp: KL + ghi chú
    res = client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": ong["id"], "rv_quantity": "40", "rv_note": "Đo lại"}]}).json()
    assert res["saved"] == 1
    assert client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": ong["id"], "rv_labor": -5}]}).status_code == 400
    d = client.get(f"/api/mep/packages/{pid}").json()
    ong = next(i for i in d["items"] if i["id"] == ong["id"])
    assert ong["rv_amount"] == 40 * 14500 and ong["review_status"] == "Cao"
    # soát xét lại vẫn giữ dòng đã sửa
    assert client.post(f"/api/mep/packages/{pid}/run", json={}).json()["kept_edits"] == 1
    hist = client.get(f"/api/mep/packages/{pid}/history").json()
    assert any(h["field"] == "rv_quantity" for h in hist)
    # xuất file gốc: giữ sheet/công thức/ô gộp, thêm cột soát xét
    out = client.get(f"/api/mep/packages/{pid}/export-original")
    assert out.status_code == 200
    wb = openpyxl.load_workbook(io.BytesIO(out.content))
    assert wb.sheetnames == ["Bia", "BG"] and wb["Bia"]["A2"].value == "=1+1"
    ws = wb["BG"]
    assert ws["H5"].value == "=D5*SUM(E5:G5)"
    assert "E3:G3" in {str(m) for m in ws.merged_cells.ranges}
    heads = [ws.cell(row=4, column=c).value for c in range(1, ws.max_column + 1)]
    assert "KL SOÁT XÉT" in heads and "GHI CHÚ SOÁT XÉT" in heads
    kl_col = heads.index("KL SOÁT XÉT") + 1
    assert ws.cell(row=6, column=kl_col).value in (40, "=ROUND(40,0)")


def test_library_keeps_sheet_structure_and_exports_original():
    data = _xlsx({
        "Vật tư": [["THƯ VIỆN"], [], ["STT", "Mã", "Tên vật tư", "ĐVT", "Đơn giá VT", "Đơn giá NC", "Nguồn"],
                   ["I", None, "CÁP ĐIỆN", None, None, None, None],
                   [1, "CD-11", "Cáp CV 1x2.5", "m", 5000, 2000, "BG A"]],
        "Nhân công": [["Mã", "Công việc", "ĐVT", "Đơn giá"], ["NC-11", "Nhân công điện", "công", 450000]],
        "Ca máy": [["Mã", "Tên máy", "ĐVT", "Đơn giá ca"], ["CM-11", "Máy khoan", "ca", 285000]],
    })
    j = client.post("/api/mep/library/versions/import", files={"file": ("tv.xlsx", data)}).json()
    vid = j["version"]["id"]
    d = client.get(f"/api/mep/library/versions/{vid}/items").json()
    assert [s["name"] for s in d["sheets"]] == ["Vật tư", "Nhân công", "Ca máy"]
    assert d["sheets"][0]["headers"] == ["STT", "Mã", "Tên vật tư", "ĐVT", "Đơn giá VT", "Đơn giá NC", "Nguồn"]
    assert [(r["row"], r["kind"]) for r in d["sheet"]["rows"]] == [(4, "group"), (5, "item")]
    item = d["sheet"]["rows"][1]["item"]
    client.put(f"/api/mep/library/items/{item['id']}", json={"labor_price": 2500})
    out = client.get(f"/api/mep/library/versions/{vid}/export")
    wb = openpyxl.load_workbook(io.BytesIO(out.content))
    assert wb.sheetnames == ["Vật tư", "Nhân công", "Ca máy"]
    ws = wb["Vật tư"]
    assert ws["A1"].value == "THƯ VIỆN" and ws["C4"].value == "CÁP ĐIỆN" and ws["F5"].value == 2500 and ws["G5"].value == "BG A"


def test_review_all_sheets_with_structure():
    wb = openpyxl.Workbook()
    wb.active.title = "Bia"
    for name, rows in {"Dien": [["STT", "Hạng mục", "ĐVT", "Khối lượng", "Đơn giá"], ["I", "CÁP", None, None, None],
                                [1, "Cáp điện CV 3x4mm2", "m", 10, 30000], [None, "Cộng", None, None, None]],
                       "Nuoc": [["TT", "Nội dung công việc", "Đvt", "SL", "Đơn giá VL", "Đơn giá NC"], [1, "Ống PPR D25", "m", 5, 40000, 15000]]}.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    pr = client.post("/api/mep/projects", json={"name": "Dự án nhiều sheet"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "code": "MS-01", "name": "Nhiều sheet", "start": "true"},
                      files={"file": ("ntp.xlsx", buf.getvalue())}).json()["id"]
    d = client.get(f"/api/mep/packages/{pid}").json()
    assert [c["sheet"] for c in d["sheet_configs"]] == ["Dien", "Nuoc"]
    assert [(r["row"], r["kind"]) for r in d["structure"]["Dien"]] == [(2, "group"), (3, "item"), (4, "group")]
    assert d["structure"]["Nuoc"][0]["cells"][1] == "Ống PPR D25"
    assert client.post(f"/api/mep/packages/{pid}/run", json={"remove_sheet": "Nuoc"}).status_code == 200
    assert [c["sheet"] for c in client.get(f"/api/mep/packages/{pid}").json()["sheet_configs"]] == ["Dien"]
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    assert "KL SOÁT XÉT" in [c.value for c in out["Dien"][1]]


def test_review_depreciation_percent():
    """Cột '% KH VẬT TƯ' là tỷ lệ khấu hao: Đơn giá = VT × %KH + NC."""
    data = _xlsx({"Bien phap": [
        ["STT", "Hạng mục", "ĐVT", "Khối lượng", "Đơn giá VT", "% KH VẬT TƯ SỐ THÁNG: 12 tháng", "Đơn giá NC", "Đơn giá", "Thành tiền"],
        [1, "Giàn giáo thép", "bộ", 10, 14950, 0.2, None, 2990, 29900],
        [2, "Tôn che chắn", "m2", 5, 82400, 0.62, 16500, 67588, 337940],
    ]})
    pr = client.post("/api/mep/projects", json={"name": "Dự án KH"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Biện pháp", "start": "true"},
                      files={"file": ("bp.xlsx", data)}).json()["id"]
    # Không làm tròn: so khớp đúng số NTP
    assert client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [], "round_unit": 0}).status_code == 200
    d = client.get(f"/api/mep/packages/{pid}").json()
    assert d["sheet_configs"][0]["mapping"]["depreciation_pct"] == 6
    it = {i["boq_description"]: i for i in d["items"]}
    gg, ton = it["Giàn giáo thép"], it["Tôn che chắn"]
    assert gg["boq_dep_pct"] == 0.2 and gg["rv_dep_pct"] == 0.2
    assert round(gg["rv_unit_total"]) == 2990 and round(ton["rv_unit_total"]) == 67588
    assert abs(gg["diff_amount"]) < 1 and gg["review_status"] == "Hợp lý"
    # Nhập "10%" → 0,1 → đơn giá 1.495
    assert client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": gg["id"], "rv_dep_pct": "10%"}]}).status_code == 200
    gg = next(i for i in client.get(f"/api/mep/packages/{pid}").json()["items"] if i["id"] == gg["id"])
    assert gg["rv_dep_pct"] == 0.1 and round(gg["rv_unit_total"]) == 1495
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    heads = [c.value for c in out["Bien phap"][1]]
    assert "% KH VT SOÁT XÉT" in heads
    dc = heads.index("% KH VT SOÁT XÉT") + 1
    assert out["Bien phap"].cell(row=2, column=dc).value == 0.1
    f = out["Bien phap"].cell(row=2, column=dc + 1).value
    assert f.startswith("=") and "*" in f, f


def test_custom_review_columns():
    pr = client.post("/api/mep/projects", json={"name": "Dự án cột"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Cột thêm", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    cols = client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [
        {"label": "Giá NCC 2", "type": "number"}, {"label": "Nhận xét", "type": "text"}]}).json()["columns"]
    assert [c["label"] for c in cols] == ["Giá NCC 2", "Nhận xét"]
    assert client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [{"label": ""}]}).status_code == 400
    d = client.get(f"/api/mep/packages/{pid}").json()
    assert len(d["custom_columns"]) == 2
    it = d["items"][0]
    n, t = cols[0]["id"], cols[1]["id"]
    assert client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "x_" + n: "abc"}]}).status_code == 400
    assert client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "x_" + n: "1.250.000", "x_" + t: "OK"}]}).json()["changes"] == 2
    it2 = next(i for i in client.get(f"/api/mep/packages/{pid}").json()["items"] if i["id"] == it["id"])
    assert json.loads(it2["rv_extra"]) == {n: 1250000, t: "OK"}
    hist = client.get(f"/api/mep/packages/{pid}/history").json()
    assert {"Giá NCC 2", "Nhận xét"} <= {h["field_label"] for h in hist}
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    vals = [c.value for ws in out.worksheets for row in ws.iter_rows() for c in row]
    assert "GIÁ NCC 2" in vals and "OK" in vals


def test_p2_column_order():
    pr = client.post("/api/mep/projects", json={"name": "Dự án thứ tự"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Thứ tự", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    assert client.get(f"/api/mep/packages/{pid}").json()["p2_order"][:4] == ["qty", "vt", "nc", "dep"]
    r = client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [{"label": "Giá NCC", "type": "number"}],
                                                          "order": ["new:0", "note", "qty", "vt", "bogus"]}).json()
    xid = "x_" + r["columns"][0]["id"]
    order = client.get(f"/api/mep/packages/{pid}").json()["p2_order"]
    assert order[:4] == [xid, "note", "qty", "vt"] and "bogus" not in order and len(order) == 11
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    for ws in out.worksheets:
        heads = [c.value for row in ws.iter_rows(max_row=6) for c in row]
        if "GIÁ NCC" in heads:
            assert heads.index("GIÁ NCC") < heads.index("GHI CHÚ SOÁT XÉT") < heads.index("KL SOÁT XÉT")
            break
    else:
        raise AssertionError("không thấy cột Giá NCC")


def test_formula_column_export():
    pr = client.post("/api/mep/projects", json={"name": "Dự án công thức"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Công thức", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    assert client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [{"label": "VAT", "type": "formula"}]}).status_code == 400
    cols = client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [
        {"label": "Thuế VAT", "type": "formula", "formula": "=ROUND([Thành tiền]*10%;0)"}]}).json()["columns"]
    assert cols[0]["formula"].startswith("=")
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    found = [c.value for ws in out.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=ROUND(")]
    assert any("*10%" in f for f in found)
    # Đơn giá mặc định làm tròn trăm: công thức ROUND(…,-2)
    assert any(f.endswith(",-2)") for f in found)


def test_add_insert_rows():
    pr = client.post("/api/mep/projects", json={"name": "Dự án thêm dòng"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Thêm dòng", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    d = client.get(f"/api/mep/packages/{pid}").json()
    sheet = d["sheet_configs"][0]["sheet"]
    rows = d["structure"][sheet]
    first, last = rows[0], rows[-1]
    a = client.post(f"/api/mep/packages/{pid}/rows", json={"anchor_sid": first["sid"], "where": "after"}).json()["id"]
    b = client.post(f"/api/mep/packages/{pid}/rows", json={"anchor_sid": first["sid"], "where": "before"}).json()["id"]
    order = [r["item_id"] for r in client.get(f"/api/mep/packages/{pid}").json()["structure"][sheet]]
    assert order.index(b) < order.index(first["item_id"]) < order.index(a) < order.index(rows[1]["item_id"])
    r = client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": a, "add_desc": "Phụ kiện bổ sung", "add_unit": "bộ",
                                                                      "rv_quantity": 2, "rv_material": 1000, "rv_labor": 500}]})
    assert r.status_code == 200
    d = client.get(f"/api/mep/packages/{pid}").json()
    it = next(i for i in d["items"] if i["id"] == a)
    assert it["boq_description"] == "Phụ kiện bổ sung" and it["rv_amount"] == 3000
    # dòng gốc không xóa được
    assert client.delete(f"/api/mep/packages/{pid}/rows/{first['item_id']}").status_code == 400
    # soát xét lại vẫn giữ dòng thêm
    client.post(f"/api/mep/packages/{pid}/run", json={})
    d = client.get(f"/api/mep/packages/{pid}").json()
    assert any(i["boq_description"] == "Phụ kiện bổ sung" and i["is_added"] for i in d["items"])
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    vals = [c.value for ws in out.worksheets for row in ws.iter_rows() for c in row]
    assert "CÁC DÒNG KSCP BỔ SUNG" in vals and "Phụ kiện bổ sung" in vals
    nid = next(i["id"] for i in d["items"] if i["is_added"] and i["boq_description"] == "")
    assert client.delete(f"/api/mep/packages/{pid}/rows/{nid}").status_code == 200


def test_cell_formula_saved():
    pr = client.post("/api/mep/projects", json={"name": "Dự án ô công thức"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Ô công thức", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    it = client.get(f"/api/mep/packages/{pid}").json()["items"][0]
    assert client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_material": 1100, "f_rv_material": "=1000*1,1"}]}).status_code == 200
    it2 = next(i for i in client.get(f"/api/mep/packages/{pid}").json()["items"] if i["id"] == it["id"])
    assert it2["rv_material"] == 1100 and json.loads(it2["rv_extra"])["f_rv_material"] == "=1000*1,1"
    assert client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "f_rv_material": "abc"}]}).status_code == 400


def test_round_unit_default_100():
    pr = client.post("/api/mep/projects", json={"name": "Dự án làm tròn"}).json()
    data = _xlsx({"BP": [["STT", "Hạng mục", "ĐVT", "Khối lượng", "Đơn giá VT", "% KH VẬT TƯ", "Đơn giá NC", "Đơn giá"],
                         [1, "Giàn giáo", "bộ", 10, 14950, 0.2, None, 2990]]})
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Làm tròn", "start": "true"},
                      files={"file": ("bp.xlsx", data)}).json()["id"]
    d = client.get(f"/api/mep/packages/{pid}").json()
    assert d["round_unit"] == 100
    it = d["items"][0]
    assert it["rv_unit_total"] == 3000 and it["rv_amount"] == 30000  # 15.000 × 20% = 3.000
    client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_material": 14950, "f_rv_material": "=E2*1"}]})
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    vals = [c.value for row in out["BP"].iter_rows() for c in row if isinstance(c.value, str)]
    assert "=ROUND(E2*1,-2)" in vals


def test_edit_ntp_values():
    pr = client.post("/api/mep/projects", json={"name": "Dự án sửa NTP"}).json()
    data = _xlsx({"BP": [["STT", "Hạng mục", "ĐVT", "Khối lượng", "Đơn giá VT", "% KH VẬT TƯ", "Đơn giá NC", "Đơn giá"],
                         [1, "Giàn giáo", "bộ", 10, 14950, 0.2, None, "=E2*F2+G2"]]})
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Sửa NTP", "start": "true"},
                      files={"file": ("bp.xlsx", data)}).json()["id"]
    client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [], "round_unit": 0})
    it = client.get(f"/api/mep/packages/{pid}").json()["items"][0]
    r = client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "boq_labor": 1000, "f_boq_labor": "=500*2", "boq_quantity": 12}]})
    assert r.status_code == 200
    it = client.get(f"/api/mep/packages/{pid}").json()["items"][0]
    assert it["boq_price"] == 14950 * 0.2 + 1000 and it["boq_quantity"] == 12
    ex = json.loads(it["rv_extra"])
    assert ex["o_boq_labor"] is None and ex["o_boq_quantity"] == 10
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    ws = out["BP"]
    assert ws["G2"].value == "=500*2" and ws["D2"].value == 12 and ws["H2"].value == "=E2*F2+G2"


def test_save_versions_files():
    pr = client.post("/api/mep/projects", json={"name": "Dự án bản lưu"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Bản lưu", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    it = client.get(f"/api/mep/packages/{pid}").json()["items"][0]
    assert client.get(f"/api/mep/packages/{pid}/versions").json() == []
    client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_quantity": 7}]})
    client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_quantity": 9}]})
    client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_quantity": 9}]})  # không đổi → không tạo bản
    import time
    for _ in range(50):  # file Excel của bản lưu được xuất nền
        vs = client.get(f"/api/mep/packages/{pid}/versions").json()
        if all(v["has_file"] for v in vs):
            break
        time.sleep(0.1)
    assert [v["no"] for v in vs] == [2, 1] and all(v["has_file"] for v in vs)
    v1 = client.get(f"/api/mep/packages/{pid}/versions/{vs[1]['id']}").json()
    assert v1["data"][str(it["id"])]["rv_quantity"] == 7
    f = client.get(f"/api/mep/packages/{pid}/versions/{vs[1]['id']}/file")
    assert f.status_code == 200 and f.content[:2] == b"PK"


def test_cell_style_saved_and_exported():
    pr = client.post("/api/mep/projects", json={"name": "Dự án định dạng"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Định dạng", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    it = client.get(f"/api/mep/packages/{pid}").json()["items"][0]
    st = {"p2:vt": {"b": True, "bg": "#ffff00", "al": "c", "bad": 1}, "p2:amt": {"nf": "d2", "c": "zzz"}}
    assert client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "st": st}]}).status_code == 200
    it = client.get(f"/api/mep/packages/{pid}").json()["items"][0]
    assert json.loads(it["rv_extra"])["st"] == {"p2:vt": {"b": True, "bg": "#FFFF00", "al": "c"}, "p2:amt": {"nf": "d2"}}
    out = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/mep/packages/{pid}/export-original").content))
    found = [c for ws in out.worksheets for row in ws.iter_rows() for c in row if c.fill and c.fill.start_color and c.fill.start_color.rgb == "00FFFF00"]
    assert found and found[0].font.bold


def test_reset_to_original():
    pr = client.post("/api/mep/projects", json={"name": "Dự án reset"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "Reset", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    d = client.get(f"/api/mep/packages/{pid}").json()
    it = d["items"][0]
    client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_quantity": 99, "boq_quantity": 77}]})
    sid = d["structure"][d["sheet_configs"][0]["sheet"]][0]["sid"]
    client.post(f"/api/mep/packages/{pid}/rows", json={"anchor_sid": sid})
    client.post(f"/api/mep/packages/{pid}/run", json={})  # giữ chỉnh sửa
    d2 = client.get(f"/api/mep/packages/{pid}").json()
    assert any(i["rv_quantity"] == 99 for i in d2["items"]) and any(i.get("is_added") for i in d2["items"])
    client.post(f"/api/mep/packages/{pid}/run", json={"reset": True})  # về hồ sơ gốc
    d3 = client.get(f"/api/mep/packages/{pid}").json()
    assert not any(i["rv_quantity"] == 99 for i in d3["items"]) and not any(i.get("is_added") for i in d3["items"])
    assert all(i["boq_quantity"] != 77 for i in d3["items"])
    assert client.get(f"/api/mep/packages/{pid}/versions").json()  # bản lưu cũ vẫn còn


def test_total_override_keeps_parts():
    pr = client.post("/api/mep/projects", json={"name": "Dự án ĐG tổng"}).json()
    pid = client.post("/api/mep/packages", data={"project_id": pr["id"], "name": "ĐG tổng", "start": "true"},
                      files={"file": ("ntp.xlsx", _ntp())}).json()["id"]
    client.put(f"/api/mep/packages/{pid}/columns", json={"columns": [], "round_unit": 0})
    it = next(i for i in client.get(f"/api/mep/packages/{pid}").json()["items"] if i.get("rv_material") is not None)
    mat = it["rv_material"]
    client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_total": 12345, "f_rv_total": "=12345"}]})
    it2 = next(i for i in client.get(f"/api/mep/packages/{pid}").json()["items"] if i["id"] == it["id"])
    assert it2["rv_material"] == mat and it2["rv_unit_total"] == 12345
    client.post(f"/api/mep/packages/{pid}/save", json={"rows": [{"id": it["id"], "rv_total": None, "f_rv_total": None}]})
    it3 = next(i for i in client.get(f"/api/mep/packages/{pid}").json()["items"] if i["id"] == it["id"])
    assert it3["rv_unit_total"] != 12345 and it3["rv_material"] == mat
