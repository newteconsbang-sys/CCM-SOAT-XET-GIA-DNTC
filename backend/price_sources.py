"""Nguồn giá cho nút "Chạy" trên bảng soát xét: lấy giá theo từng loại thư viện đã nạp.

- dntc  : Đơn giá ĐNTC chuẩn (bản đang áp dụng / mới nhất)
- ncc   : Báo giá NCC – so đầu mục, áp giá NCC
- hd    : Hợp đồng tham khảo
- default (Mặc định): lấy ĐNTC chuẩn làm cơ sở; Báo giá NCC nạp SAU ĐNTC chuẩn thì ghi đè;
          dòng vẫn chưa có giá thì lấy Hợp đồng tham khảo (nạp trước hay sau đều được)
- truc  : Đơn giá nhân công trực (bảng tra cứu) – khớp theo nội dung công việc
- kh    : %KH theo thời gian (bảng Khấu hao) – xem review_sheets.kh_infos
"""
import json
import re
from typing import Any, Dict, List, Optional

from backend.mep_engine import match_single_item


def latest_version(conn, kind: str) -> Optional[Dict[str, Any]]:
    order = "(status='active') DESC, id DESC" if kind == "dntc" else "id DESC"
    r = conn.execute(f"SELECT id, name, created_at, COALESCE(kind,'dntc') AS kind, note FROM mep_library_versions "
                     f"WHERE COALESCE(kind,'dntc')=? AND status<>'reference' ORDER BY {order} LIMIT 1", (kind,)).fetchone()
    return dict(r) if r else None


def _match_version(conn, items: List[Dict[str, Any]], vid: int, threshold: float) -> Dict[int, Dict[str, Any]]:
    lib = [dict(r) for r in conn.execute("SELECT * FROM mep_libraries WHERE version=?", (vid,))]
    if not lib:
        return {}
    by_id = {x["id"]: x for x in lib}
    out: Dict[int, Dict[str, Any]] = {}
    for it in items:
        q = {"code": it.get("boq_code"), "description": it.get("boq_description"), "specification": it.get("boq_spec"),
             "brand": it.get("boq_brand"), "model": it.get("boq_model"), "unit": it.get("boq_unit"), "price": it.get("boq_price")}
        if not (q["description"] or "").strip():
            continue
        try:
            m = match_single_item(q, lib, [], threshold)
        except Exception:
            continue
        row = by_id.get(m.get("lib_id")) if m.get("lib_id") else None
        if not row or not float(row.get("total_price") or 0):
            continue
        out[it["id"]] = {"material": row.get("material_price"), "labor": row.get("labor_price"),
                         "machine": row.get("machine_price"), "total": float(row.get("total_price") or 0),
                         "lib_desc": row.get("description"), "score": m.get("match_score")}
    return out


def _norm(s: Any) -> str:
    from backend.depreciation import norm
    return norm(s)


def _truc_prices(conn, items: List[Dict[str, Any]]) -> Dict[int, Dict[str, Any]]:
    """Bảng nhân công trực: dòng có mô tả + số tiền → khớp mô tả hạng mục (trùng từ nhiều nhất)."""
    v = latest_version(conn, "truc")
    if not v:
        return {}
    rows: List[tuple] = []
    for sh in conn.execute("SELECT * FROM mep_library_sheets WHERE version_id=?", (v["id"],)):
        heads = [_norm(h) for h in json.loads(sh["headers_json"] or "[]")]
        di = next((i for i, h in enumerate(heads) if h.startswith(("thoi gian", "noi dung", "mo ta", "ten", "dien giai"))), 1)
        pi = next((i for i, h in enumerate(heads) if "tien" in h or "don gia" in h or "gia" == h), None)
        for r in conn.execute("SELECT cells_json FROM mep_library_rows WHERE sheet_id=? AND kind<>'title'", (sh["id"],)):
            c = json.loads(r["cells_json"] or "[]")
            desc = c[di] if di < len(c) else None
            val = c[pi] if pi is not None and pi < len(c) else None
            if pi is None:  # không rõ cột tiền: lấy số lớn nhất trong dòng
                nums = [x for x in c if isinstance(x, (int, float))]
                val = max(nums) if nums else None
            if desc and isinstance(val, (int, float)) and val > 0:
                rows.append((set(re.findall(r"[a-z0-9]+", _norm(desc))), float(val), str(desc)))
    out: Dict[int, Dict[str, Any]] = {}
    for it in items:
        words = set(re.findall(r"[a-z0-9]+", _norm(it.get("boq_description"))))
        if not words or not ({"truc", "tang"} & words):
            continue
        best = max(rows, key=lambda r: (len(words & r[0]) / max(1, len(r[0] | words))), default=None)
        if best and len(words & best[0]) / max(1, len(best[0] | words)) >= 0.35:
            out[it["id"]] = {"material": None, "labor": best[1], "machine": None, "total": best[1], "lib_desc": best[2]}
    return out


def version_by_id(conn, vid: int) -> Optional[Dict[str, Any]]:
    r = conn.execute("SELECT id, name, created_at, COALESCE(kind,'dntc') AS kind, note FROM mep_library_versions WHERE id=?", (vid,)).fetchone()
    return dict(r) if r else None


def prices(conn, items: List[Dict[str, Any]], source: str, threshold: float = 10, version_id: Optional[int] = None) -> Dict[str, Any]:
    """{"prices": {item_id: {material, labor, machine, total, src}}, "used": [tên bản], "note": …}"""
    used: List[str] = []
    res: Dict[int, Dict[str, Any]] = {}

    def take(kind: str, label: str, overwrite: bool, only_missing: bool = False):
        v = (version_by_id(conn, version_id) if version_id and source == kind else None) or latest_version(conn, kind)
        if not v:
            return None
        got = _match_version(conn, [i for i in items if not (only_missing and i["id"] in res)], v["id"], threshold)
        for k, x in got.items():
            if overwrite or k not in res:
                res[k] = {**x, "src": f"{label} {v['name']}"}
        fn = (v.get("note") or "").replace("Nạp từ ", "") if v.get("note") else ""
        used.append(f"{label} {fn or v['name']} ({len(got)})")
        return v

    if source == "truc":
        for k, x in _truc_prices(conn, items).items():
            res[k] = {**x, "src": "Nhân công trực"}
        v = latest_version(conn, "truc")
        used.append(f"Đơn giá nhân công trực {v['name']} ({len(res)})" if v else "chưa nạp Đơn giá nhân công trực")
    elif source in ("dntc", "ncc", "hd"):
        lbl = {"dntc": "ĐNTC chuẩn", "ncc": "Báo giá NCC", "hd": "Hợp đồng tham khảo"}[source]
        if not take(source, lbl, True):
            used.append(f"chưa nạp {lbl}")
    else:  # default
        base = take("dntc", "ĐNTC chuẩn", True)
        ncc = latest_version(conn, "ncc")
        if ncc and (not base or str(ncc["created_at"] or "") > str(base["created_at"] or "")):
            take("ncc", "Báo giá NCC", True)
        elif ncc:
            used.append(f"bỏ qua Báo giá NCC {ncc['name']} (nạp trước ĐNTC chuẩn)")
        take("hd", "Hợp đồng tham khảo", False, only_missing=True)
    return {"prices": {str(k): v for k, v in res.items()}, "used": used}
