"""Soát xét giá theo gói thầu: đề xuất giá KSCP–MEP từ thư viện, tính đơn giá tổng / thành tiền /
chênh lệch, lưu chỉnh sửa trực tiếp kèm lịch sử và giữ chỉnh sửa khi soát xét lại."""
import json
import math
import re
import uuid
from typing import Any, Dict, List, Optional

from backend.mep_db import get_connection, now_iso
from backend.xlsx_utils import parse_number

RV_FIELDS = ("rv_quantity", "rv_material", "rv_labor", "rv_machine", "rv_total", "rv_note", "rv_dep_pct")
NUM_FIELDS = ("rv_quantity", "rv_material", "rv_labor", "rv_machine", "rv_total", "rv_dep_pct")
FIELD_LABELS = {"rv_quantity": "Khối lượng", "rv_material": "Đơn giá VT", "rv_labor": "Đơn giá NC",
                "rv_machine": "Khấu hao", "rv_total": "Đơn giá tổng", "rv_note": "Ghi chú", "rv_dep_pct": "% KH vật tư",
                "boq_stt": "STT (dòng thêm)", "boq_description": "Hạng mục (dòng thêm)", "boq_unit": "ĐVT (dòng thêm)",
                "x_ov_ntp_amt": "Thành tiền NTP", "x_ov_rv_amt": "Thành tiền soát xét"}
# Kết quả so NTP: NTP báo cao hơn soát xét → "Cao"; thấp hơn → "Thấp"; bằng → "Hợp lý"
STATUS_OK, STATUS_DOWN, STATUS_ADJ = "Hợp lý", "Cao", "Thấp"
EPS = 0.5  # đồng


def ensure_schema():
    with get_connection() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS mep_review_edits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            package_id INTEGER,
            session_id TEXT,
            item_id INTEGER,
            item_label TEXT,
            field TEXT,
            old_value TEXT,
            new_value TEXT,
            changed_by TEXT,
            changed_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_rev_edits_pkg ON mep_review_edits(package_id);
        """)
        cols = {r[1] for r in conn.execute("PRAGMA table_info(mep_packages)")}
        for name, typ in (("header_row", "INTEGER"), ("mapping_json", "TEXT")):
            if name not in cols:
                conn.execute(f"ALTER TABLE mep_packages ADD COLUMN {name} {typ}")


def round_unit(v: Optional[float], unit: float) -> Optional[float]:
    """Làm tròn như Excel ROUND về bội của unit (100 → …00); unit 0 = không làm tròn."""
    if v is None or not unit:
        return v
    v = float(v)
    return math.copysign(math.floor(abs(v) / unit + 0.5 + 1e-9) * unit, v)


def unit_total(material: Optional[float], labor: Optional[float], machine: Optional[float], total: Optional[float],
               dep: Optional[float] = None, unit: float = 0) -> float:
    """Đơn giá tổng = VT (× %KH nếu có) + NC + Máy; không tách thành phần thì dùng đơn giá tổng.
    unit > 0: từng đơn giá và đơn giá tổng làm tròn về bội unit (mặc định 100đ)."""
    material, labor, machine = (round_unit(x, unit) for x in (material, labor, machine))
    if dep is not None and material is not None:
        material = float(material) * float(dep)
    parts = [p for p in (material, labor, machine) if p is not None]
    if parts:
        return float(round_unit(sum(parts), unit))
    return float(round_unit(float(total or 0), unit))


# Thành tiền gõ / công thức riêng trên ô (như ô Excel độc lập): lưu trong rv_extra
OV_KEYS = ("ov_ntp_amt", "ov_rv_amt")


def _ov(it: Dict[str, Any], key: str) -> Optional[float]:
    ex = it.get("_ex")
    if ex is None:
        try:
            ex = json.loads(it.get("rv_extra") or "{}") if isinstance(it.get("rv_extra"), str) else (it.get("rv_extra") or {})
        except ValueError:
            ex = {}
    v = ex.get(key) if isinstance(ex, dict) else None
    return float(v) if isinstance(v, (int, float)) else None


def ntp_values(it: Dict[str, Any]) -> Dict[str, Any]:
    qty = float(it.get("boq_quantity") or 0)
    price = float(it.get("boq_price") or 0)
    ov = _ov(it, "ov_ntp_amt")
    return {"quantity": qty, "material": it.get("boq_material"), "labor": it.get("boq_labor"),
            "machine": it.get("boq_machine"), "total": price, "amount": ov if ov is not None else qty * price}


def propose(it: Dict[str, Any]) -> Dict[str, Any]:
    """Đề xuất giá soát xét ban đầu. NTP cao hơn thư viện (vượt ngưỡng) → hạ từng thành phần về mức thư viện;
    còn lại giữ nguyên giá NTP. Khối lượng giữ theo NTP."""
    rv = {"rv_quantity": it.get("boq_quantity") or 0, "rv_material": it.get("boq_material"),
          "rv_labor": it.get("boq_labor"), "rv_machine": it.get("boq_machine"),
          "rv_total": None, "rv_note": None, "rv_dep_pct": it.get("boq_dep_pct")}
    if it.get("boq_dep_done"):  # vật tư đã khấu hao một phần (dùng lại) → không tính nhân công lắp
        rv["rv_labor"] = None
    has_parts = any(it.get(k) is not None for k in ("boq_material", "boq_labor", "boq_machine"))
    if it.get("status") == "higher" and it.get("lib_price"):
        lib = {"rv_material": it.get("lib_material"), "rv_labor": it.get("lib_labor"), "rv_machine": it.get("lib_machine")}
        if has_parts and any(v is not None for v in lib.values()):
            for k, src in (("rv_material", "boq_material"), ("rv_labor", "boq_labor"), ("rv_machine", "boq_machine")):
                ntp_v, lib_v = it.get(src), lib[k]
                if ntp_v is not None and lib_v is not None:
                    rv[k] = min(float(ntp_v), float(lib_v))
        else:
            rv.update({"rv_material": None, "rv_labor": None, "rv_machine": None})
            rv["rv_total"] = min(float(it.get("boq_price") or 0), float(it["lib_price"]))
        rv["rv_note"] = None
    if not has_parts and rv["rv_total"] is None:
        rv["rv_total"] = float(it.get("boq_price") or 0)
    return rv


def review_values(it: Dict[str, Any]) -> Dict[str, Any]:
    qty = float(it["rv_quantity"]) if it.get("rv_quantity") is not None else float(it.get("boq_quantity") or 0)
    if it.get("_ru"):  # có làm tròn: khối lượng soát xét làm tròn số nguyên (ROUND(…,0))
        qty = round_unit(qty, 1)
    dep = it.get("rv_dep_pct") if it.get("rv_dep_pct") is not None else it.get("boq_dep_pct")
    if it.get("rv_total") is not None:  # ĐG tổng nhập riêng được ưu tiên
        total = float(round_unit(float(it["rv_total"]), it.get("_ru") or 0))
    else:
        lab = None if it.get("boq_dep_done") else it.get("rv_labor")  # có % ĐÃ KH → ĐG NC soát xét bị khóa (= 0)
        total = unit_total(it.get("rv_material"), lab, it.get("rv_machine"), it.get("boq_price"), dep, it.get("_ru") or 0)
    ov = _ov(it, "ov_rv_amt")
    return {"quantity": qty, "total": total, "amount": ov if ov is not None else qty * total}


# PA3 – Phần 3 (P.KSCP soát xét): bộ cột soát xét thứ 2, lưu trong rv_extra "k3_*"; ô chưa sửa = theo Phần 2
K3_NUM = ("k3_qty", "k3_vt", "k3_nc", "k3_dep", "k3_total", "k3_amt")
K3_TEXT = ("k3_note",)


def _extra(it: Dict[str, Any]) -> Dict[str, Any]:
    ex = it.get("rv_extra")
    if isinstance(ex, dict):
        return ex
    try:
        return json.loads(ex or "{}")
    except ValueError:
        return {}


def k3_values(it: Dict[str, Any]) -> Dict[str, Any]:
    ex, r, ru = _extra(it), review_values(it), it.get("_ru") or 0
    g = lambda k: float(ex[k]) if isinstance(ex.get(k), (int, float)) else None
    qty = g("k3_qty")
    qty = r["quantity"] if qty is None else (round_unit(qty, 1) if ru else qty)
    vt = g("k3_vt") if g("k3_vt") is not None else it.get("rv_material")
    nc = None if it.get("boq_dep_done") else (g("k3_nc") if g("k3_nc") is not None else it.get("rv_labor"))
    dep = g("k3_dep") if g("k3_dep") is not None else (it.get("rv_dep_pct") if it.get("rv_dep_pct") is not None else it.get("boq_dep_pct"))
    parts = any(v is not None for v in (vt, nc, it.get("rv_machine")))
    if g("k3_total") is not None:
        total = float(round_unit(g("k3_total"), ru))
    elif parts and (it.get("rv_total") is None or any(g(k) is not None for k in ("k3_vt", "k3_nc", "k3_dep"))):
        total = unit_total(vt, nc, it.get("rv_machine"), it.get("boq_price"), dep, ru)
    else:
        total = r["total"]
    amt = g("k3_amt") if g("k3_amt") is not None else qty * total
    n = ntp_values(it)["amount"]
    st = STATUS_DOWN if amt < n - EPS else STATUS_ADJ if amt > n + EPS else STATUS_OK
    return {"qty": qty, "vt": vt, "nc": nc, "dep": dep, "total": total, "amount": amt, "diff": amt - n, "status": st,
            "note": ex.get("k3_note")}


def totals_k3(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    t = totals(items)
    rev = 0.0
    counts: Dict[str, int] = {STATUS_OK: 0, STATUS_DOWN: 0, STATUS_ADJ: 0}
    for i in items:
        k = k3_values(i)
        rev += k["amount"]
        counts[k["status"]] = counts.get(k["status"], 0) + 1
    t.update(review=rev, diff=rev - t["ntp"], counts=counts,
             diff_percent=round((rev - t["ntp"]) / t["ntp"] * 100, 2) if t["ntp"] else None)
    return t


def status_of(it: Dict[str, Any]) -> str:
    n, r = ntp_values(it), review_values(it)
    if r["amount"] < n["amount"] - EPS:
        return STATUS_DOWN
    if r["amount"] > n["amount"] + EPS:
        return STATUS_ADJ
    return STATUS_OK


def enrich(it: Dict[str, Any]) -> Dict[str, Any]:
    n, r = ntp_values(it), review_values(it)
    it["ntp_total"], it["ntp_amount"] = n["total"], n["amount"]
    it["rv_unit_total"], it["rv_amount"] = r["total"], r["amount"]
    it["diff_amount"] = r["amount"] - n["amount"]
    it["review_status"] = status_of(it)
    it["review_price"] = r["total"]  # tương thích màn hình/báo cáo cũ
    lib = it.get("lib_price")
    # Dòng có %KH: so thư viện theo giá trị đầy đủ VT + NC (+ máy) chứ không theo giá đã khấu hao
    ntp_cmp = n["total"]
    if it.get("boq_dep_pct") is not None and it.get("boq_material") is not None:
        ntp_cmp = sum(float(it.get(k) or 0) for k in ("boq_material", "boq_labor", "boq_machine"))
    it["lib_diff_percent"] = round((ntp_cmp - lib) / lib * 100, 1) if lib else None
    st = it.get("status")
    it["lib_flag"] = ("none" if not lib else "uncertain" if st in ("manual_check", "spec_mismatch")
                      else "higher" if st == "higher" else "lower" if st == "lower" else "ok")
    return it


def session_items(conn, session_id: Optional[str]) -> List[Dict[str, Any]]:
    if not session_id:
        return []
    rows = conn.execute(
        """SELECT i.*, l.material_price AS lib_material, l.labor_price AS lib_labor,
                  l.machine_price AS lib_machine, l.waste_percent AS lib_waste, l.item_type AS lib_type
           FROM mep_review_items i LEFT JOIN mep_libraries l ON l.id = i.lib_id
           WHERE i.session_id = ? ORDER BY i.excel_row""", (session_id,)).fetchall()
    pk = conn.execute("SELECT round_unit FROM mep_packages WHERE session_id=?", (session_id,)).fetchone()
    ru = float(pk["round_unit"]) if pk and pk["round_unit"] is not None else 0
    out = []
    for r in rows:
        d = dict(r)
        d.pop("candidates_json", None)
        d["_ru"] = ru
        out.append(enrich(d))
    return out


CUSTOM_TYPES = {"number": "Số", "text": "Chữ", "formula": "Công thức"}


def custom_columns(conn, package_id: int) -> List[Dict[str, Any]]:
    row = conn.execute("SELECT custom_cols_json FROM mep_packages WHERE id=?", (package_id,)).fetchone()
    try:
        return json.loads(row["custom_cols_json"] or "[]") if row else []
    except (TypeError, ValueError):
        return []


def clean_custom_columns(cols: Any) -> List[Dict[str, Any]]:
    """Kiểm tra danh sách cột tự tạo: [{id, label, type}] – id giữ nguyên để không mất dữ liệu đã nhập."""
    if not isinstance(cols, list) or len(cols) > 30:
        raise ValueError("Danh sách cột không hợp lệ (tối đa 30 cột)")
    out, seen = [], set()
    for c in cols:
        label = str((c or {}).get("label") or "").strip()[:60]
        if not label:
            raise ValueError("Tên cột không được để trống")
        typ = (c or {}).get("type") if (c or {}).get("type") in CUSTOM_TYPES else "text"
        cid = re.sub(r"[^a-z0-9]", "", str((c or {}).get("id") or "").lower())[:12] or uuid.uuid4().hex[:8]
        if cid in seen:
            cid = uuid.uuid4().hex[:8]
        seen.add(cid)
        col = {"id": cid, "label": label, "type": typ}
        if typ == "formula":
            f = str((c or {}).get("formula") or "").strip()[:300]
            if not f:
                raise ValueError(f"Cột {label}: chưa nhập công thức")
            col["formula"] = f if f.startswith("=") else "=" + f
        out.append(col)
    return out


P2_KEYS = ["qty", "vt", "nc", "dep", "total", "amt", "diff", "pct", "status", "note"]


def p2_order(conn, package_id: int, cols: Optional[List[Dict[str, Any]]] = None) -> List[str]:
    """Thứ tự cột Phần 2: cột chuẩn + cột tự tạo ("x_<id>"); thiếu cột nào thì bổ sung theo mặc định."""
    cols = custom_columns(conn, package_id) if cols is None else cols
    row = conn.execute("SELECT p2_order_json FROM mep_packages WHERE id=?", (package_id,)).fetchone()
    try:
        saved = json.loads(row["p2_order_json"] or "[]") if row else []
    except (TypeError, ValueError):
        saved = []
    return normalize_order(saved, cols)


def normalize_order(order: Any, cols: List[Dict[str, Any]]) -> List[str]:
    valid = P2_KEYS + ["x_" + c["id"] for c in cols]
    out = [("pct" if k == "lib" else k) for k in (order if isinstance(order, list) else [])]  # bản cũ: Giá TV → % chênh lệch
    out = [k for k in out if k in valid]
    out = list(dict.fromkeys(out))
    for k in valid:
        if k not in out:
            out.append(k)
    return out


def init_session(session_id: str, previous_session: Optional[str] = None) -> Dict[str, int]:
    """Gán giá đề xuất cho phiên mới; dòng đã sửa tay ở phiên trước (cùng dòng Excel + cùng hạng mục) được giữ lại."""
    kept = 0
    with get_connection() as conn:
        old: Dict[tuple, Dict[str, Any]] = {}
        if previous_session:
            for r in conn.execute("SELECT * FROM mep_review_items WHERE session_id=? AND rv_touched=1", (previous_session,)):
                old[(r["sheet_name"], r["excel_row"], (r["boq_description"] or "").strip().lower())] = dict(r)
        items = [i for i in session_items(conn, session_id) if not i.get("is_added")]
        for it in items:
            prev = old.get((it["sheet_name"], it["excel_row"], (it["boq_description"] or "").strip().lower()))
            if prev:
                rv = {k: prev[k] for k in RV_FIELDS}
                conn.execute("UPDATE mep_review_items SET rv_extra=? WHERE id=?", (prev.get("rv_extra"), it["id"]))
                touched = 1
                kept += 1
            else:
                rv = propose(it)
                touched = 0
            conn.execute(f"UPDATE mep_review_items SET {', '.join(k + '=?' for k in RV_FIELDS)}, rv_touched=? WHERE id=?",
                         (*[rv[k] for k in RV_FIELDS], touched, it["id"]))
            if prev:  # chép cả lịch sử của dòng sang dòng mới
                conn.execute("UPDATE mep_review_edits SET item_id=? WHERE item_id=?", (it["id"], prev["id"]))
    return {"items": len(items), "kept_edits": kept}


# Số liệu NTP sửa được trên bảng (Phần 1): giữ số gốc trong rv_extra "o_<cột>" để đối chiếu
NTP_FIELDS = ("boq_quantity", "boq_material", "boq_labor", "boq_dep_pct", "boq_price", "boq_dep_done")
FIELD_LABELS.update({"boq_quantity": "Khối lượng NTP", "boq_material": "Đơn giá VT NTP", "boq_labor": "Đơn giá NC NTP",
                     "boq_dep_pct": "% KH VT NTP", "boq_price": "Đơn giá NTP", "boq_dep_done": "% đã KH",
                     "x_d_start": "Thời gian bắt đầu", "x_d_end": "Thời gian kết thúc"})
DATE_KEYS = ("d_start", "d_end")  # ngày bắt đầu / kết thúc sửa trên bảng (rv_extra)
TEXT_KEYS = ("t_stt", "t_desc")   # STT / Mô tả công việc sửa trên bảng (dòng của file; rv_extra)
FIELD_LABELS.update({"x_t_stt": "STT", "x_t_desc": "Mô tả công việc"})
FIELD_LABELS.update({"x_k3_qty": "KL P.KSCP", "x_k3_vt": "ĐG VT P.KSCP", "x_k3_nc": "ĐG NC P.KSCP", "x_k3_dep": "% KH P.KSCP",
                     "x_k3_total": "ĐG tổng P.KSCP", "x_k3_amt": "Thành tiền P.KSCP", "x_k3_note": "Ghi chú P.KSCP"})


def _parse_patch(row: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for k in NTP_FIELDS:
        if k not in row:
            continue
        v = row[k]
        if v in (None, ""):
            out[k] = 0.0 if k in ("boq_quantity", "boq_price") else None
            continue
        n = parse_number(v)
        if n is None or n < 0:
            raise ValueError(f"{FIELD_LABELS[k]} phải là số ≥ 0")
        out[k] = n / 100 if k in ("boq_dep_pct", "boq_dep_done") and n > 1 else n
    for k in RV_FIELDS:
        if k not in row:
            continue
        v = row[k]
        if k == "rv_note":
            out[k] = (str(v).strip() or None) if v is not None else None
            continue
        if v in (None, ""):
            out[k] = None
            continue
        n = parse_number(v)
        if n is None:
            raise ValueError(f"{FIELD_LABELS[k]} phải là số")
        if k == "rv_dep_pct" and n > 1:
            n = n / 100  # nhập 20 hoặc 20% → 0,2
        if n < 0:
            raise ValueError(f"{FIELD_LABELS[k]} không được âm")
        out[k] = n
    return out


STYLE_KEYS = {"b": bool, "i": bool, "u": bool, "c": str, "bg": str, "al": str, "nf": str}


def _clean_style(st: Any) -> Dict[str, Any]:
    """{cột: {b,i,u,c,bg,al,nf}} – chỉ giữ thuộc tính hợp lệ."""
    out: Dict[str, Any] = {}
    if not isinstance(st, dict):
        return out
    for col, sty in list(st.items())[:80]:
        if not isinstance(sty, dict) or len(str(col)) > 40:
            continue
        cs = {}
        for k, typ in STYLE_KEYS.items():
            v = sty.get(k)
            if v in (None, "", False):
                continue
            if typ is bool:
                cs[k] = True
            elif k in ("c", "bg") and re.fullmatch(r"#[0-9a-fA-F]{6}", str(v)):
                cs[k] = str(v).upper()
            elif k == "al" and v in ("l", "c", "r"):
                cs[k] = v
            elif k == "nf" and v in ("int", "d2", "pct"):
                cs[k] = v
        if cs:
            out[str(col)] = cs
    return out


def save_rows(package_id: int, session_id: str, rows: List[Dict[str, Any]], user: str = "") -> Dict[str, Any]:
    if not isinstance(rows, list):
        raise ValueError("Dữ liệu dòng không hợp lệ")
    saved, changes = 0, 0
    with get_connection() as conn:
        xcols = {c["id"]: c for c in custom_columns(conn, package_id)}
        for row in rows:
            try:
                iid = int(row.get("id"))
            except (TypeError, ValueError):
                raise ValueError("Thiếu mã dòng")
            cur = conn.execute("SELECT * FROM mep_review_items WHERE id=? AND session_id=?", (iid, session_id)).fetchone()
            if not cur:
                continue
            cur = dict(cur)
            try:
                patch = _parse_patch(row)
            except ValueError as e:
                raise ValueError(f"Dòng {cur.get('boq_stt') or cur['excel_row']}: {e}")
            # Dòng KSCP thêm vào: sửa được STT / Hạng mục / ĐVT
            if cur.get("is_added"):
                for k, col in (("add_stt", "boq_stt"), ("add_desc", "boq_description"), ("add_unit", "boq_unit")):
                    if k in row:
                        patch[col] = (str(row[k] or "").strip()[:500]) or ("" if col == "boq_description" else None)
            new = {**cur, **patch}
            # Sửa VT / NC / %KH của NTP → đơn giá NTP tính lại = VT × %KH + NC (+ máy), trừ khi gõ thẳng đơn giá
            if any(k in patch for k in ("boq_material", "boq_labor", "boq_dep_pct")) and "boq_price" not in patch \
                    and (new.get("boq_material") is not None or new.get("boq_labor") is not None):
                mat = float(new.get("boq_material") or 0) * (float(new["boq_dep_pct"]) if new.get("boq_dep_pct") is not None else 1)
                new["boq_price"] = mat + float(new.get("boq_labor") or 0) + float(new.get("boq_machine") or 0)
            # ĐG tổng có giá trị riêng (gõ / công thức trên ô ĐG tổng) thì ưu tiên – không xóa ĐG VT / ĐG NC
            diff = {k: new.get(k) for k in (*RV_FIELDS, *NTP_FIELDS, "boq_stt", "boq_description", "boq_unit") if not _same(cur.get(k), new.get(k))}
            # Cột tự tạo: khóa "x_<id>" trong dòng gửi lên
            try:
                cur_x = json.loads(cur.get("rv_extra") or "{}")
            except ValueError:
                cur_x = {}
            xdiff: Dict[str, Any] = {}
            for k, v in row.items():
                # Công thức của ô (gõ = trên bảng): lưu chữ công thức cạnh giá trị, khóa "f_<cột>"
                if str(k).startswith("f_") and (k[2:] in NUM_FIELDS or k[2:] in NTP_FIELDS or (k[2:4] == "x_" and (k[4:] in xcols or k[4:] in OV_KEYS or k[4:] in K3_NUM))):
                    fv = (str(v).strip()[:300] or None) if v is not None else None
                    if fv and not fv.startswith("="):
                        raise ValueError(f"Dòng {cur.get('boq_stt') or cur['excel_row']}: công thức phải bắt đầu bằng =")
                    if not _same(cur_x.get(k), fv):
                        xdiff[k] = fv
                    continue
                if str(k).startswith("x_") and k[2:] in TEXT_KEYS:  # STT / mô tả sửa trên bảng
                    tv = (str(v).strip()[:500] or None) if v is not None else None
                    if not _same(cur_x.get(k[2:]), tv):
                        xdiff[k[2:]] = tv
                    continue
                if str(k).startswith("x_") and k[2:] in DATE_KEYS:  # ngày bắt đầu / kết thúc sửa trên bảng
                    from backend.depreciation import parse_date
                    dv = None
                    if v not in (None, ""):
                        d = parse_date(v)
                        if not d:
                            raise ValueError(f"Dòng {cur.get('boq_stt') or cur['excel_row']}: ngày phải dạng dd/mm/yyyy")
                        dv = d.strftime("%d/%m/%Y")
                    if not _same(cur_x.get(k[2:]), dv):
                        xdiff[k[2:]] = dv
                    continue
                if str(k).startswith("x_") and k[2:] in OV_KEYS:  # Thành tiền NTP / soát xét gõ riêng
                    nv = parse_number(v) if v not in (None, "") else None
                    if v not in (None, "") and nv is None:
                        raise ValueError(f"Dòng {cur.get('boq_stt') or cur['excel_row']}: Thành tiền phải là số")
                    if not _same(cur_x.get(k[2:]), nv):
                        xdiff[k[2:]] = nv
                    continue
                if str(k).startswith("x_") and k[2:] in K3_NUM + K3_TEXT:  # PA3 – Phần 3 P.KSCP soát xét
                    if k[2:] in K3_TEXT:
                        nv = (str(v).strip()[:500] or None) if v is not None else None
                    else:
                        nv = parse_number(v) if v not in (None, "") else None
                        if v not in (None, "") and (nv is None or nv < 0):
                            raise ValueError(f"Dòng {cur.get('boq_stt') or cur['excel_row']}: Phần 3 phải là số ≥ 0")
                        if k[2:] == "k3_dep" and nv is not None and nv > 1:
                            nv = nv / 100
                    if not _same(cur_x.get(k[2:]), nv):
                        xdiff[k[2:]] = nv
                    continue
                if not str(k).startswith("x_") or k[2:] not in xcols:
                    continue
                col = xcols[k[2:]]
                if col["type"] == "formula":  # cột công thức: tự tính, không nhập
                    continue
                if v in (None, ""):
                    nv = None
                elif col["type"] == "number":
                    nv = parse_number(v)
                    if nv is None:
                        raise ValueError(f"Dòng {cur.get('boq_stt') or cur['excel_row']}: {col['label']} phải là số")
                else:
                    nv = str(v).strip()[:500] or None
                if not _same(cur_x.get(col["id"]), nv):
                    xdiff[col["id"]] = nv
            if "st" in row:  # định dạng ô (đậm, nghiêng, màu, căn lề, định dạng số) theo cột
                st = _clean_style(row.get("st"))
                if not _same(json.dumps(cur_x.get("st") or {}, sort_keys=True), json.dumps(st, sort_keys=True)):
                    xdiff["st"] = st or None
            if not diff and not xdiff:
                continue
            label = f"{cur.get('boq_stt') or ''} {cur['boq_description']}".strip()
            for cid, v in list(xdiff.items()):
                if cid.startswith(("f_", "o_")) or cid == "st":  # công thức: không ghi lịch sử riêng (giá trị ô đã ghi)
                    if v is None:
                        cur_x.pop(cid, None)
                    else:
                        cur_x[cid] = v
                    continue
                conn.execute("INSERT INTO mep_review_edits (package_id, session_id, item_id, item_label, field, old_value, new_value, changed_by, changed_at)"
                             " VALUES (?,?,?,?,?,?,?,?,?)",
                             (package_id, session_id, iid, label, "x_" + cid, _txt(cur_x.get(cid)), _txt(v), user, now_iso()))
                changes += 1
                if v is None:
                    cur_x.pop(cid, None)
                else:
                    cur_x[cid] = v
            for k in NTP_FIELDS:  # lần đầu sửa số NTP: lưu số gốc
                if k in diff and "o_" + k not in cur_x:
                    cur_x["o_" + k] = cur.get(k)
                    xdiff["o_" + k] = cur.get(k)
            if xdiff:
                diff["rv_extra"] = json.dumps(cur_x, ensure_ascii=False)
            for k, v in diff.items():
                if k == "rv_extra":
                    continue
                conn.execute("INSERT INTO mep_review_edits (package_id, session_id, item_id, item_label, field, old_value, new_value, changed_by, changed_at)"
                             " VALUES (?,?,?,?,?,?,?,?,?)",
                             (package_id, session_id, iid, label, k, _txt(cur.get(k)), _txt(v), user, now_iso()))
                changes += 1
            rv_total_final = unit_total(new.get("rv_material"), new.get("rv_labor"), new.get("rv_machine"),
                                        new.get("rv_total") if new.get("rv_total") is not None else new.get("boq_price"),
                                        new.get("rv_dep_pct") if new.get("rv_dep_pct") is not None else new.get("boq_dep_pct"))
            action = "kept_boq" if abs(rv_total_final - float(cur.get("boq_price") or 0)) < EPS else \
                "applied_lib" if cur.get("lib_price") and abs(rv_total_final - float(cur["lib_price"])) < EPS else "edited_custom"
            conn.execute(f"UPDATE mep_review_items SET {', '.join(k + '=?' for k in diff)}, rv_touched=1, edited_price=?, user_action=?,"
                         " user_note=?, warning_color='green', updated_at=? WHERE id=?",
                         (*diff.values(), rv_total_final, action, new.get("rv_note"), now_iso(), iid))
            saved += 1
        if saved:
            conn.execute("UPDATE mep_packages SET updated_at=? WHERE id=?", (now_iso(), package_id))
    return {"saved": saved, "changes": changes}


def history(package_id: int, limit: int = 500) -> List[Dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute("SELECT * FROM mep_review_edits WHERE package_id=? ORDER BY id DESC LIMIT ?", (package_id, limit)).fetchall()
        xl = {"x_" + c["id"]: c["label"] for c in custom_columns(conn, package_id)}
    out = []
    for r in rows:
        d = dict(r)
        f = d.get("field") or ""
        d["field_label"] = FIELD_LABELS.get(f) or xl.get(f) or ("Cột đã xóa" if f.startswith("x_") else f)
        out.append(d)
    return out


def log_event(package_id: int, session_id: Optional[str], text: str, user: str = ""):
    with get_connection() as conn:
        conn.execute("INSERT INTO mep_review_edits (package_id, session_id, field, new_value, changed_by, changed_at) VALUES (?,?,?,?,?,?)",
                     (package_id, session_id, "event", text, user, now_iso()))


def totals(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    ntp = sum(i["ntp_amount"] for i in items)
    rev = sum(i["rv_amount"] for i in items)
    counts: Dict[str, int] = {STATUS_OK: 0, STATUS_DOWN: 0, STATUS_ADJ: 0}
    for i in items:
        counts[i["review_status"]] = counts.get(i["review_status"], 0) + 1
    return {"ntp": ntp, "review": rev, "diff": rev - ntp,
            "diff_percent": round((rev - ntp) / ntp * 100, 2) if ntp else None, "counts": counts,
            "not_in_library": sum(1 for i in items if not i.get("lib_price"))}


def _same(a: Any, b: Any) -> bool:
    if isinstance(a, (int, float)) or isinstance(b, (int, float)):
        if a is None or b is None:
            return a is None and b is None
        return abs(float(a) - float(b)) < 1e-9
    return (a or None) == (b or None)


def _txt(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)
