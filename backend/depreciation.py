"""Khấu hao vật tư (%KH): nhận biết hạng mục nào là vật tư có khấu hao và thuộc nhóm nào trong bảng
"KH VẬT TƯ" (T=1…T=42 tháng), học tên vật tư từ hồ sơ NTP đã soát xét.

- CATS: danh mục nhóm khấu hao đúng như bảng KH VẬT TƯ 2026 (Xa biển / Gần biển).
- classify(tên, nhóm cha): luật theo từ khóa (cáp theo vật liệu / số lõi / tiết diện, đèn theo công suất,
  các bộ phận lắp trong tủ điện → Tủ thi công…), ưu tiên tên đã học (khớp nguyên tên).
- learn_from_boq(file): đọc hồ sơ BOQ có cột "% KH VẬT TƯ"; dòng có %KH < 100% là vật tư khấu hao → ghi nhớ tên.
"""
import json
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[1]
SEED_FILE = Path(__file__).resolve().parent / "kh_seed.json"      # tên đã học sẵn (đi kèm tool)
LEARN_FILE = ROOT / "data" / "kh_learned.json"                     # tên học thêm trên máy người dùng

G_CU, G_AL, G_DIEN, G_NUOC = "CÁP ĐIỆN ĐỒNG", "CÁP ĐIỆN NHÔM", "VẬT TƯ THIẾT BỊ ĐIỆN", "VẬT TƯ THIẾT BỊ NƯỚC"
CABLE_NAMES = ["Cáp 10 đến 240mm2/1c", "Cáp 6.0 đến 8.0 mm2/1c", "Cáp 1.5 đến 5.5mm2/1c",
               "Cáp 10 đến 240mm2/2c,3c,4c,5c", "Cáp 5.5 đến 8.0mm2/2c,3c,4c,5c", "Cáp 1.5 đến 4.0mm2/2c,3c,4c,5c",
               "Cáp đồng trần từ ≥16mm2 phần nổi", "Cáp đồng trần âm đất", "Cáp đồng trần đến ≤10mm2"]
DIEN_NAMES = ["Tủ thi công", "Ổ cắm thông dụng", "Ổ cắm công nghiệp di động", "Jắc cắm công nghiệp", "Điện kế",
              "Đèn huỳnh quang", "Đèn Metal Halide", "Đèn led cao áp 150W bảo hành 2 năm", "Đèn pha led 250W bảo hành 2 năm",
              "Đèn pha led 400W bảo hành 2 năm", "Đèn báo không lưu", "Quạt hút âm trần, tường.", "Quạt công nghiệp",
              "Máy lạnh và ống đồng", "Ống sắt D60x6m (chống trụ)", "Cọc đồng tiếp địa 16x2,4m", "Kim thu sét"]
NUOC_NAMES = ["Máy bơm chìm (không bao gồm hệ thống bơm thoát nước bề mặt)", "Máy bơm tăng áp", "Máy bơm hỏa tiễn",
              "Máy bơm rửa xe", "Tủ bơm biến tần", "Tủ bơm nước", "Bồn áp lực", "Bồn nước PVC 500L -2000L",
              "Bồn nước trên 2000L", "Thiết bị vệ sinh và phụ kiện", "Công tắc phao", "Ống và phụ kiện ống cấp thoát nước"]
CATS: List[Tuple[str, str]] = ([(G_CU, n) for n in CABLE_NAMES] + [(G_AL, n) for n in CABLE_NAMES]
                               + [(G_DIEN, n) for n in DIEN_NAMES] + [(G_NUOC, n) for n in NUOC_NAMES])


def norm(s: Any) -> str:
    t = unicodedata.normalize("NFD", str(s or "")).replace("đ", "d").replace("Đ", "D")
    t = re.sub(r"[̀-ͯ]", "", t).lower().replace("²", "2")
    return re.sub(r"\s+", " ", t).strip()


def cat_id(group: str, name: str) -> str:
    return f"{norm(group)}|{norm(name)}"


CAT_BY_ID = {cat_id(g, n): (g, n) for g, n in CATS}

TU_PARTS = ("den bao pha", "cau chi", "mccb", "mcb", "cau dao", "shuntrip", "shunt trip", "ro le", "role ", "relay", "thanh cai",
            "cable clamp", "vo tu", "gia do tu", "mat che", "teminal", "terminal", "elcb", "rccb", "contactor", "khoi dong tu",
            "dong ho a", "dong ho v", "ampe", "volt", "bien dong", "bien dong")


def _cable(t: str) -> Optional[Tuple[str, str]]:
    if "clamp" in t or not (t.startswith("cap") or "xlpe" in t or re.search(r"\b(cu|al)/", t) or t.startswith("day cv")
                            or t.startswith("day dien")):
        return None
    grp = G_AL if ("nhom" in t or re.search(r"\bal/", t)) else G_CU
    if "tran" in t:
        sizes = [float(x.replace(",", ".")) for x in re.findall(r"(\d+(?:[.,]\d+)?)\s*mm", t)] or [0]
        if "am dat" in t or "chon" in t:
            return grp, CABLE_NAMES[7]
        return grp, CABLE_NAMES[6] if max(sizes) >= 16 else CABLE_NAMES[8]
    m = re.search(r"(\d+)\s*c\b", t) or re.search(r"\b(\d)\s*[x×]\s*\d", t)
    cores = int(m.group(1)) if m else 1
    sizes = [float(x.replace(",", ".")) for x in re.findall(r"[x×+]\s*(\d+(?:[.,]\d+)?)", t)]
    sizes = [s for s in sizes if s < 1000] or [float(x.replace(",", ".")) for x in re.findall(r"(\d+(?:[.,]\d+)?)\s*mm", t)] or [0]
    s = max(sizes)
    if cores <= 1:
        name = CABLE_NAMES[0] if s >= 10 else CABLE_NAMES[1] if s >= 6 else CABLE_NAMES[2]
    else:
        name = CABLE_NAMES[3] if s >= 10 else CABLE_NAMES[4] if s >= 5.5 else CABLE_NAMES[5]
    return grp, name


def classify(desc: str, parent: str = "") -> Optional[Tuple[str, str]]:
    """Nhóm khấu hao (nhóm, tên dòng trong bảng KH) hoặc None nếu không phải vật tư khấu hao."""
    t, p = norm(desc), norm(parent)
    if not t:
        return None
    learned = learned_map().get(t)
    if learned and learned in CAT_BY_ID:
        return CAT_BY_ID[learned]
    # Tủ điều khiển bơm
    if t.startswith("tu") and "bom" in t:
        return G_NUOC, NUOC_NAMES[4] if "bien tan" in t else NUOC_NAMES[5]
    # Bộ phận lắp trong tủ điện / dòng nằm dưới hạng mục "Tủ …" → Tủ thi công
    in_tu = bool(re.match(r"^tu\b", p)) and "bom" not in p
    if any(k in t for k in TU_PARTS) or (in_tu and not t.startswith(("den led", "den pha", "bom"))) or t.startswith("tu dien") \
            or t.startswith("tu thi cong") or (t.startswith("tu ") and "bom" not in t):
        if not t.startswith("den bao khong luu"):
            return G_DIEN, DIEN_NAMES[0]
    c = _cable(t)
    if c:
        return c
    if t.startswith("den") or " den " in f" {t} ":
        if "khong luu" in t:
            return G_DIEN, DIEN_NAMES[10]
        if "metal" in t or "halide" in t:
            return G_DIEN, DIEN_NAMES[6]
        if "huynh quang" in t or "tuyp" in t or "tube" in t:
            return G_DIEN, DIEN_NAMES[5]
        w = [int(x) for x in re.findall(r"(\d+)\s*w\b", t)]
        watt = max(w) if w else 0
        if "pha" in t or watt > 150:
            return G_DIEN, DIEN_NAMES[8] if watt <= 250 else DIEN_NAMES[9]
        return G_DIEN, DIEN_NAMES[7]
    if "o cam" in t:
        return G_DIEN, DIEN_NAMES[2] if ("cong nghiep" in t or "di dong" in t) else DIEN_NAMES[1]
    if t.startswith(("jac", "jack", "phich")):
        return G_DIEN, DIEN_NAMES[3]
    if "dien ke" in t or "cong to" in t:
        return G_DIEN, DIEN_NAMES[4]
    if t.startswith("quat"):
        return G_DIEN, DIEN_NAMES[11] if "hut" in t else DIEN_NAMES[12]
    if "may lanh" in t or "dieu hoa" in t:
        return G_DIEN, DIEN_NAMES[13]
    if "tru dien" in t or "chong tru" in t or t.startswith("ong sat"):
        return G_DIEN, DIEN_NAMES[14]
    if "coc" in t and ("tiep dia" in t or "dong" in t):
        return G_DIEN, DIEN_NAMES[15]
    if "kim thu set" in t:
        return G_DIEN, DIEN_NAMES[16]
    if "bom" in t and not t.startswith(("ong", "van", "day")):
        if "chim" in t:
            return G_NUOC, NUOC_NAMES[0]
        if "hoa tien" in t:
            return G_NUOC, NUOC_NAMES[2]
        if "rua xe" in t:
            return G_NUOC, NUOC_NAMES[3]
        return G_NUOC, NUOC_NAMES[1]
    if ("binh" in t or "bon" in t) and "ap luc" in t:
        return G_NUOC, NUOC_NAMES[6]
    if t.startswith(("bon nuoc", "bon chua", "tec nuoc")):
        v = [int(x) for x in re.findall(r"(\d+)\s*(?:l|lit)\b", t)]
        return G_NUOC, NUOC_NAMES[8] if v and max(v) > 2000 else NUOC_NAMES[7]
    if "thiet bi ve sinh" in t or t.startswith(("lavabo", "bon cau", "xi bet", "voi ")):
        return G_NUOC, NUOC_NAMES[9]
    if "phao" in t:
        return G_NUOC, NUOC_NAMES[10]
    if t.startswith("ong") and any(k in t for k in ("ppr", "upvc", "pvc", "hdpe", "cap nuoc", "thoat nuoc")):
        return G_NUOC, NUOC_NAMES[11]
    return None


# ------------------------------------------------------------------ tên đã học

_cache: Dict[str, Any] = {}


def _load(path: Path) -> Dict[str, Dict[str, Any]]:
    try:
        return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    except Exception:
        return {}


def learned_all() -> Dict[str, Dict[str, Any]]:
    key = (SEED_FILE.stat().st_mtime if SEED_FILE.exists() else 0, LEARN_FILE.stat().st_mtime if LEARN_FILE.exists() else 0)
    if _cache.get("key") != key:
        d = _load(SEED_FILE)
        d.update(_load(LEARN_FILE))
        _cache.update(key=key, data=d)
    return _cache["data"]


def learned_map() -> Dict[str, str]:
    return {k: v["cat"] for k, v in learned_all().items() if v.get("cat")}  # (dự phòng: tên gắn nhóm KH cụ thể)


def _boq_rows(path: Path) -> List[Dict[str, Any]]:
    """Dòng hạng mục của BOQ: tên, ĐVT, %KH, nhóm cha gần nhất (dòng không có ĐVT)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out: List[Dict[str, Any]] = []
    try:
        for ws in wb.worksheets:
            if getattr(ws, "sheet_state", "visible") != "visible":
                continue
            rows = list(ws.iter_rows(values_only=True))
            hdr, cols = None, {}
            for i, r in enumerate(rows[:20]):
                hs = [norm(v) for v in r]
                d = next((j for j, h in enumerate(hs) if h.startswith(("mo ta", "ten vat tu", "noi dung", "dien giai", "ten cong viec"))), None)
                k = next((j for j, h in enumerate(hs) if "% kh" in h or h.startswith("%kh") or "khau hao" in h), None)
                if d is not None and k is not None:
                    u = next((j for j, h in enumerate(hs) if h in ("dvt", "dv", "don vi", "don vi tinh") or h.startswith("don vi")), None)
                    hdr, cols = i, {"desc": d, "kh": k, "unit": u}
                    break
            if hdr is None:
                continue
            parent = ""
            for r in rows[hdr + 1:]:
                desc = r[cols["desc"]] if cols["desc"] < len(r) else None
                if desc is None or not str(desc).strip():
                    continue
                unit = r[cols["unit"]] if cols["unit"] is not None and cols["unit"] < len(r) else None
                kh = r[cols["kh"]] if cols["kh"] < len(r) else None
                if not unit or not str(unit).strip():
                    parent = str(desc).strip()
                    continue
                out.append({"desc": str(desc).strip(), "unit": str(unit).strip(), "kh": kh if isinstance(kh, (int, float)) else None,
                            "parent": parent})
    finally:
        wb.close()
    return out


# Nhóm trong bảng KH luôn 100% (dùng hết, không thu hồi) → mặc định khấu hao 100%
FULL_100 = {cat_id(g, n) for g, n in CATS if n in (CABLE_NAMES[2], CABLE_NAMES[5], CABLE_NAMES[7], CABLE_NAMES[8], DIEN_NAMES[5],
                                                   DIEN_NAMES[15], NUOC_NAMES[8], NUOC_NAMES[9], NUOC_NAMES[10], NUOC_NAMES[11])}


def kh_status(desc: str, parent: str = "") -> Dict[str, Any]:
    """Vật tư này có thể khấu hao dưới 100% không. Ưu tiên tên đã học từ hồ sơ; không có thì suy theo nhóm vật tư.
    → {"can_kh": bool, "src": "learned" | "rule" | "default", "group": tên nhóm KH hoặc ""}"""
    v = learned_all().get(norm(desc))
    cat = classify(desc, parent)
    label = f"{cat[0]} › {cat[1]}" if cat else ""
    if v and "kh" in v:
        return {"can_kh": v["kh"] == "lt100", "src": "learned", "group": label}
    if cat:
        return {"can_kh": cat_id(*cat) not in FULL_100, "src": "rule", "group": label}
    return {"can_kh": False, "src": "default", "group": ""}


def learn_from_boq(path: Path, source: str = "") -> Dict[str, Any]:
    """Học từ hồ sơ BOQ có cột "% KH VẬT TƯ": tên có %KH < 100% → vật tư có thể khấu hao dưới 100%;
    tên chỉ gặp với %KH = 100% → mặc định 100%. Ghi vào data/kh_learned.json (chỉ lưu tên + loại, không lưu số %)."""
    rows = _boq_rows(path)
    data = _load(LEARN_FILE)
    seed = _load(SEED_FILE)
    n_lt, n_full = 0, 0
    for r in rows:
        if r["kh"] is None or r["kh"] <= 0:
            continue
        key = norm(r["desc"])
        kind = "lt100" if r["kh"] < 1 else "100"
        old = data.get(key) or seed.get(key)
        if old and (old.get("kh") == "lt100" or old.get("kh") == kind):
            continue  # đã biết; "có khấu hao < 100%" không bị ghi đè bởi 100%
        data[key] = {"name": r["desc"].split("\n")[0].strip(), "kh": kind, "src": source}
        if kind == "lt100":
            n_lt += 1
        else:
            n_full += 1
    LEARN_FILE.parent.mkdir(parents=True, exist_ok=True)
    LEARN_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    _cache.clear()
    return {"rows": len(rows), "added_lt100": n_lt, "added_100": n_full}


def summary() -> Dict[str, Any]:
    """Danh sách tên đã học: có thể khấu hao < 100% / mặc định 100%."""
    lt, full = set(), set()
    for v in learned_all().values():
        (lt if v.get("kh") == "lt100" else full).add(v.get("name", ""))
    return {"lt100": sorted(lt), "full100": sorted(full)}


# ------------------------------------------------------------------ %KH theo thời gian (bảng KH VẬT TƯ)
import calendar
import datetime as _dt
import math


def parse_date(v: Any) -> Optional[_dt.date]:
    """Ô ngày: kiểu ngày Excel, 'dd/mm/yyyy' (ngày vượt cuối tháng như 31/06 → cuối tháng), 'yyyy-mm-dd'."""
    if v is None or v == "":
        return None
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return _dt.date(1899, 12, 30) + _dt.timedelta(days=int(v))
    t = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$", t)
        if not m:
            return None
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        y += 2000 if y < 100 else 0
    if not (1 <= mo <= 12):
        return None
    d = max(1, min(d, calendar.monthrange(y, mo)[1]))
    return _dt.date(y, mo, d)


def months_between(start: Any, end: Any) -> Optional[int]:
    """Số tháng khấu hao = ROUNDUP((Kết thúc − Bắt đầu)/30, 0), tối thiểu 1."""
    a, b = parse_date(start), parse_date(end)
    if not a or not b or b < a:
        return None
    return max(1, math.ceil((b - a).days / 30 - 1e-9))


def region_of_sheet(name: str) -> str:
    n = norm(name)
    return "gan" if "gan bien" in n else "xa" if "xa bien" in n else ""


def load_tables(conn) -> Dict[str, Any]:
    """Bảng KH VẬT TƯ từ thư viện Khấu hao mới nhất: {"version": tên, "sheets": {sheet: {"region", "cats": {cat_id: [T1..]}}}}."""
    v = conn.execute("SELECT id, name FROM mep_library_versions WHERE kind='kh' ORDER BY id DESC LIMIT 1").fetchone()
    if not v:
        return {"version": None, "sheets": {}}
    out: Dict[str, Any] = {}
    for sh in conn.execute("SELECT * FROM mep_library_sheets WHERE version_id=? ORDER BY sheet_index", (v["id"],)):
        heads = [norm(h) for h in json.loads(sh["headers_json"] or "[]")]
        tcol = {}
        for i, h in enumerate(heads):
            m = re.match(r"^t\s*=\s*(\d+)$", h)
            if m:
                tcol[int(m.group(1))] = i
        if not tcol:
            continue
        name_i = next((i for i, h in enumerate(heads) if h.startswith(("ten", "noi dung", "mo ta"))), 1)
        nmax = max(tcol)
        cats: Dict[str, List[Optional[float]]] = {}
        group = ""
        for r in conn.execute("SELECT cells_json, kind FROM mep_library_rows WHERE sheet_id=? ORDER BY seq, id", (sh["id"],)):
            if r["kind"] == "title":
                continue
            cells = json.loads(r["cells_json"] or "[]")
            nm = cells[name_i] if name_i < len(cells) else None
            if not nm:
                continue
            vals = []
            for t in range(1, nmax + 1):
                x = cells[tcol[t]] if t in tcol and tcol[t] < len(cells) else None
                if isinstance(x, str):
                    s = x.strip().replace(",", ".")
                    try:
                        x = float(s[:-1]) / 100 if s.endswith("%") else float(s)
                    except ValueError:
                        x = None
                vals.append(float(x) if isinstance(x, (int, float)) else None)
            stt = str(cells[0]).strip() if cells and cells[0] is not None else ""
            unit_i = next((i for i, h in enumerate(heads) if h in ("dv", "dvt", "don vi")), None)
            no_unit = unit_i is None or unit_i >= len(cells) or not cells[unit_i]
            if all(x is None for x in vals) or (no_unit and re.fullmatch(r"[IVXLA-Z]+", stt)):
                group = str(nm)
                continue
            gk = next((g for g in (G_CU, G_AL, G_DIEN, G_NUOC) if norm(group).startswith(norm(g))), "")
            cats[cat_id(gk, str(nm))] = vals
        out[sh["sheet_name"]] = {"region": region_of_sheet(sh["sheet_name"]), "cats": cats}
    return {"version": v["name"], "sheets": out}


def pick_sheet(tables: Dict[str, Any], region: str) -> Optional[str]:
    sh = tables.get("sheets") or {}
    for n, s in sh.items():
        if s["region"] == region:
            return n
    return next(iter(sh), None)


def expected_pct(desc: str, parent: str, months: Optional[int], tables: Dict[str, Any], region: str) -> Dict[str, Any]:
    """%KH soát xét theo thời gian: vật tư có thể khấu hao < 100% → tra bảng KH tại T = số tháng; còn lại 100%."""
    st = kh_status(desc, parent)
    info = {"months": months, "can": st["can_kh"], "group": st["group"], "pct": None, "sheet": None}
    if months is None:
        return info
    if not st["can_kh"]:
        info["pct"] = 1.0
        return info
    cat = classify(desc, parent)
    sheet = pick_sheet(tables, region)
    if not cat or not sheet:
        return info
    vals = tables["sheets"][sheet]["cats"].get(cat_id(*cat))
    if not vals:
        return info
    t = min(months, len(vals))
    v = vals[t - 1]
    info.update(pct=v, sheet=sheet, t=t)
    return info
