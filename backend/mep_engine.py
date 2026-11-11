"""MEP Price Review Matching & Comparison Engine.
Strictly deterministic: Python, RapidFuzz, regex, MEP domain rules. No external LLMs.
"""
import re
import unicodedata
from typing import List, Dict, Any, Optional, Tuple
from rapidfuzz import fuzz

# Unit equivalence groups
UNIT_GROUPS = [
    {"cái", "chiếc", "cây", "bộ", "quả", "hạt", "tấm", "khẩu"},
    {"m", "mét", "md", "m dài"},
    {"m2", "m²", "mét vuông"},
    {"m3", "m³", "mét khối"},
    {"kg", "kilogram", "kí"},
    {"tấn"},
    {"cuộn", "lô"},
    {"hộp", "thùng", "gói"}
]

def are_units_compatible(u1: Optional[str], u2: Optional[str]) -> bool:
    """Checks whether two units can be compared without conversion."""
    if not u1 or not u2:
        return True # if one is missing, allow review with warning
    n1 = unicodedata.normalize('NFC', str(u1)).strip().lower()
    n2 = unicodedata.normalize('NFC', str(u2)).strip().lower()
    if n1 == n2:
        return True
    for g in UNIT_GROUPS:
        if n1 in g and n2 in g:
            return True
    return False

def normalize_mep_description(text: str) -> str:
    """Normalizes MEP material descriptions into canonical form for matching."""
    if not text:
        return ""
    t = unicodedata.normalize('NFC', str(text)).strip().lower()
    
    # 1. Superscripts & Symbols
    t = t.replace('²', '2').replace('³', '3')
    t = t.replace('ø', 'd').replace('phi ', 'd').replace('dn ', 'dn')
    t = t.replace('±', '+-')

    # 2. Electrical terms
    t = re.sub(r'\b(dây điện|dây cáp|cáp điện|dây cv|dây|cáp)\b', 'cap', t)
    t = re.sub(r'\b(cầu dao tự động|aptomat|át|cb)\b', 'mccb', t)
    t = re.sub(r'\b(bóng đèn|đèn chiếu sáng|đèn led|đèn)\b', 'den', t)
    t = re.sub(r'\b(tủ điện|tủ pccc|tủ hạ thế|tủ)\b', 'tu', t)

    # 3. Plumbing / Piping terms
    t = re.sub(r'\b(ống nhựa|ống cấp nước|ống thoát nước|ống dẫn nước|ống)\b', 'ong', t)
    t = re.sub(r'\b(van cổng|van chặn|gate valve)\b', 'van_cong', t)
    t = re.sub(r'\b(van bướm|butterfly valve)\b', 'van_buom', t)
    t = re.sub(r'\b(van một chiều|van 1 chiều|check valve)\b', 'van_1_chieu', t)
    t = re.sub(r'\b(van bi|ball valve)\b', 'van_bi', t)
    t = re.sub(r'\b(van giảm áp|prv)\b', 'van_giam_ap', t)

    # 4. HVAC terms
    t = re.sub(r'\b(điều hòa không khí|máy lạnh|điều hòa)\b', 'dieu_hoa', t)
    t = re.sub(r'\b(quạt thông gió|quạt hút|quạt)\b', 'quat', t)

    # 5. Cables cross section formatting: 4.0mm2 -> 4mm2, 1x4mm2, 3x2.5+1x1.5
    t = re.sub(r'(\d+)\s*x\s*(\d+(?:\.\d+)?)\s*(?:mm2)?', r'\1x\2mm2', t)
    t = re.sub(r'(?<![x\d])(\d+)\.0\s*mm2', r'\1mm2', t)
    t = re.sub(r'(?<![x\d])(\d+)\s*mm2', r'1x\1mm2', t)

    # 6. Pipe diameters: d 20 -> d20, dn 25 -> dn25
    t = re.sub(r'\b(d|dn)\s*(\d+)\b', r'\1\2', t)

    # 7. Ratings and values: 1600 A -> 1600a, 50 kA -> 50ka, pn 16 -> pn16
    t = re.sub(r'(\d+)\s*(a|ka|w|kw|hp|v|kva|pn|bar)\b', r'\1\2', t)

    # 8. Punctuation cleanup
    t = re.sub(r'[\-_/,:;+]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def extract_mep_specs(text: str) -> Dict[str, str]:
    """Extracts critical technical parameters to detect spec conflicts."""
    specs = {}
    if not text:
        return specs
    t = unicodedata.normalize('NFC', str(text)).lower()
    t = t.replace('²', '2').replace('³', '3').replace('ø', 'd').replace('phi ', 'd')

    # Cable size: e.g. 1x4mm2, 4x16mm2, 4.0mm2, 4mm2
    cable_m = re.search(r'(\d+\s*x\s*\d+(?:\.\d+)?(?:\s*\+\s*\d+\s*x\s*\d+(?:\.\d+)?)?\s*mm2|\b\d+(?:\.\d+)?\s*mm2)', t)
    if cable_m:
        raw_cs = re.sub(r'\s+', '', cable_m.group(0))
        # Standardize 4.0 -> 4, 4mm2 -> 1x4mm2
        raw_cs = re.sub(r'\.0(?=mm2|\+)', '', raw_cs)
        if not re.search(r'^\d+x', raw_cs):
            raw_cs = f"1x{raw_cs}"
        specs['cable_size'] = raw_cs

    # Diameter / Pipe size: e.g. dn50, d20, d=25mm
    dn_m = re.search(r'\b(?:dn|d)\s*(\d+)\b', t)
    if dn_m:
        specs['diameter'] = f"dn{dn_m.group(1)}"

    # Pressure: e.g. pn10, pn16, pn25, 10bar
    pn_m = re.search(r'\b(?:pn\s*(\d+)|(\d+)\s*bar)\b', t)
    if pn_m:
        specs['pressure'] = f"pn{pn_m.group(1) or pn_m.group(2)}"

    # Amperage: e.g. 1600a, 100a, 63a, 32a
    amp_m = re.search(r'\b(\d+)\s*a\b', t)
    if amp_m:
        specs['amperage'] = f"{amp_m.group(1)}a"

    # Breaking capacity: e.g. 50ka, 65ka, 36ka
    ka_m = re.search(r'\b(\d+)\s*ka\b', t)
    if ka_m:
        specs['breaking_capacity'] = f"{ka_m.group(1)}ka"

    # Power: e.g. 15kw, 18w, 5hp
    power_m = re.search(r'\b(\d+(?:\.\d+)?)\s*(kw|w|hp)\b', t)
    if power_m:
        specs['power'] = f"{power_m.group(1)}{power_m.group(2)}"

    # Poles: e.g. 3p, 2p, 1p, 4p
    pole_m = re.search(r'\b([1-4])\s*p\b', t)
    if pole_m:
        specs['poles'] = f"{pole_m.group(1)}p"

    return specs

def check_spec_conflict(specs1: Dict[str, str], specs2: Dict[str, str]) -> Optional[str]:
    """Returns mismatch reason string if any core technical specification conflicts."""
    for key in ['cable_size', 'diameter', 'pressure', 'amperage', 'breaking_capacity', 'poles']:
        v1 = specs1.get(key)
        v2 = specs2.get(key)
        if v1 and v2 and v1 != v2:
            return f"Lệch thông số {key}: BOQ ({v1}) khác Thư viện ({v2})"
    return None

def is_non_item_row(desc: str, unit: Optional[str], price: Optional[float], qty: Optional[float]) -> bool:
    """Detects header rows, category titles, section rows, and subtotals."""
    d = str(desc or '').strip()
    if not d:
        return True
    d_lower = d.lower()

    # Subtotals / Totals
    if re.match(r'^(tổng|cộng|tổng cộng|tổng chi phí|tổng cộng chi phí|chi phí xây dựng|thuế vat|vat|grand total|subtotal|total)\b', d_lower):
        return True

    # Roman numeral / section headers without price/qty
    if re.match(r'^(phần\s+[ivxldcm\d]+|[ivxldcm]+\.|\d+\.\s+[a-zà-ỹ\s]+$)', d_lower):
        if (price is None or price <= 0) and (qty is None or qty <= 0):
            return True

    # Pure header rows
    if d_lower in ['nội dung', 'hạng mục', 'mô tả công việc', 'tên vật tư', 'diễn giải', 'description', 'stt']:
        return True

    # If both price and quantity are missing and unit is missing, likely a category title
    if (price is None or price <= 0) and (unit is None or not str(unit).strip()):
        return True

    return False

def compare_price(boq_price: float, lib_price: float, threshold_percent: float) -> Tuple[float, float, str, str]:
    """Calculates price difference and assigns status and warning color."""
    if lib_price is None or lib_price <= 0:
        return 0.0, 0.0, 'not_found', 'yellow'
    
    diff = boq_price - lib_price
    diff_percent = (diff / lib_price) * 100.0

    if diff_percent > threshold_percent:
        status = 'higher'
        color = 'red'
    elif diff_percent < -threshold_percent:
        status = 'lower'
        color = 'yellow' # abnormally low or missing scope
    else:
        status = 'passed'
        color = 'green'

    return round(diff, 2), round(diff_percent, 2), status, color

def match_single_item(
    item: Dict[str, Any],
    library_items: List[Dict[str, Any]],
    learning_memory: List[Dict[str, Any]],
    threshold_percent: float = 10.0
) -> Dict[str, Any]:
    """Matches a single BOQ row against library using priority rules."""
    boq_code = str(item.get('code') or '').strip()
    boq_desc = str(item.get('description') or '').strip()
    boq_spec = str(item.get('specification') or '').strip()
    boq_brand = str(item.get('brand') or '').strip().lower()
    boq_model = str(item.get('model') or '').strip().lower()
    boq_unit = str(item.get('unit') or '').strip()
    boq_price = float(item.get('price') or 0.0)

    combined_boq_text = f"{boq_desc} {boq_spec} {boq_brand} {boq_model}".strip()
    norm_boq_text = normalize_mep_description(combined_boq_text)
    boq_specs_dict = extract_mep_specs(combined_boq_text)

    # PRIORITY 0: Learning Memory (user confirmed past alias)
    for mem in learning_memory:
        mem_boq_desc = mem.get('boq_description', '').strip()
        mem_boq_code = mem.get('boq_code', '').strip()
        is_mem_match = False
        if mem_boq_code and boq_code and mem_boq_code.lower() == boq_code.lower():
            is_mem_match = True
        elif mem_boq_desc and fuzz.token_sort_ratio(normalize_mep_description(mem_boq_desc), norm_boq_text) > 92:
            is_mem_match = True

        if is_mem_match:
            lib_price = mem.get('confirmed_price', 0.0)
            diff, diff_pct, status, color = compare_price(boq_price, lib_price, threshold_percent)
            return {
                "lib_id": None,
                "lib_code": mem.get('library_code'),
                "lib_description": mem.get('library_description'),
                "lib_spec": "",
                "lib_brand": "",
                "lib_model": "",
                "lib_unit": mem.get('library_unit') or boq_unit,
                "lib_price": lib_price,
                "lib_source": f"Bộ nhớ duyệt trước: {mem.get('notes') or 'Đã xác nhận'}",
                "difference": diff,
                "diff_percent": diff_pct,
                "status": status,
                "warning_color": color,
                "match_method": "memory",
                "match_score": 100.0,
                "candidates": []
            }

    # PRIORITY 1 & 2: Exact Code Match
    if boq_code:
        for lib in library_items:
            lib_code = str(lib.get('code') or '').strip()
            if lib_code and lib_code.lower() == boq_code.lower():
                # Check unit compatibility
                lib_unit = str(lib.get('unit') or '').strip()
                if are_units_compatible(boq_unit, lib_unit):
                    lib_price = float(lib.get('total_price') or 0.0)
                    diff, diff_pct, status, color = compare_price(boq_price, lib_price, threshold_percent)
                    return {
                        "lib_id": lib.get('id'),
                        "lib_code": lib_code,
                        "lib_description": lib.get('description'),
                        "lib_spec": lib.get('specification'),
                        "lib_brand": lib.get('brand'),
                        "lib_model": lib.get('model'),
                        "lib_unit": lib_unit,
                        "lib_price": lib_price,
                        "lib_source": lib.get('source') or 'Thư viện đơn giá',
                        "difference": diff,
                        "diff_percent": diff_pct,
                        "status": status,
                        "warning_color": color,
                        "match_method": "code_exact",
                        "match_score": 100.0,
                        "candidates": []
                    }

    # PRIORITY 3: Brand + Model + Unit Match
    if boq_brand and boq_model:
        for lib in library_items:
            lib_brand = str(lib.get('brand') or '').strip().lower()
            lib_model = str(lib.get('model') or '').strip().lower()
            lib_unit = str(lib.get('unit') or '').strip()
            if lib_brand == boq_brand and lib_model == boq_model and are_units_compatible(boq_unit, lib_unit):
                lib_price = float(lib.get('total_price') or 0.0)
                diff, diff_pct, status, color = compare_price(boq_price, lib_price, threshold_percent)
                return {
                    "lib_id": lib.get('id'),
                    "lib_code": lib.get('code'),
                    "lib_description": lib.get('description'),
                    "lib_spec": lib.get('specification'),
                    "lib_brand": lib.get('brand'),
                    "lib_model": lib.get('model'),
                    "lib_unit": lib_unit,
                    "lib_price": lib_price,
                    "lib_source": lib.get('source') or 'Thư viện đơn giá',
                    "difference": diff,
                    "diff_percent": diff_pct,
                    "status": status,
                    "warning_color": color,
                    "match_method": "brand_model_unit",
                    "match_score": 95.0,
                    "candidates": []
                }

    # PRIORITY 4: Normalized RapidFuzz Matching + Spec Conflict Prevention
    scored_candidates = []
    spec_conflict_candidates = []

    for lib in library_items:
        lib_unit = str(lib.get('unit') or '').strip()
        if not are_units_compatible(boq_unit, lib_unit):
            continue

        norm_lib_text = lib.get('normalized_desc') or normalize_mep_description(
            f"{lib.get('description','')} {lib.get('specification','')} {lib.get('brand','')} {lib.get('model','')}"
        )
        
        # Token sort ratio & token set ratio
        r_sort = fuzz.token_sort_ratio(norm_boq_text, norm_lib_text)
        r_set = fuzz.token_set_ratio(norm_boq_text, norm_lib_text)
        score = max(r_sort, r_set * 0.95)

        if score >= 60.0:
            lib_specs_dict = extract_mep_specs(f"{lib.get('description','')} {lib.get('specification','')}")
            conflict = check_spec_conflict(boq_specs_dict, lib_specs_dict)
            cand_info = {
                "lib_id": lib.get('id'),
                "lib_code": lib.get('code'),
                "lib_description": lib.get('description'),
                "lib_spec": lib.get('specification'),
                "lib_brand": lib.get('brand'),
                "lib_model": lib.get('model'),
                "lib_unit": lib_unit,
                "lib_price": float(lib.get('total_price') or 0.0),
                "lib_source": lib.get('source') or 'Thư viện đơn giá',
                "score": round(score, 1),
                "conflict": conflict
            }
            if conflict:
                spec_conflict_candidates.append(cand_info)
            else:
                scored_candidates.append(cand_info)

    scored_candidates.sort(key=lambda x: x['score'], reverse=True)
    spec_conflict_candidates.sort(key=lambda x: x['score'], reverse=True)

    if scored_candidates and scored_candidates[0]['score'] >= 80.0:
        best = scored_candidates[0]
        lib_price = best['lib_price']
        diff, diff_pct, status, color = compare_price(boq_price, lib_price, threshold_percent)
        return {
            "lib_id": best['lib_id'],
            "lib_code": best['lib_code'],
            "lib_description": best['lib_description'],
            "lib_spec": best['lib_spec'],
            "lib_brand": best['lib_brand'],
            "lib_model": best['lib_model'],
            "lib_unit": best['lib_unit'],
            "lib_price": lib_price,
            "lib_source": best['lib_source'],
            "difference": diff,
            "diff_percent": diff_pct,
            "status": status,
            "warning_color": color,
            "match_method": "rapidfuzz",
            "match_score": best['score'],
            "candidates": scored_candidates[:5]
        }
    
    if scored_candidates and scored_candidates[0]['score'] >= 65.0:
        best = scored_candidates[0]
        lib_price = best['lib_price']
        diff, diff_pct, _, _ = compare_price(boq_price, lib_price, threshold_percent)
        return {
            "lib_id": best['lib_id'],
            "lib_code": best['lib_code'],
            "lib_description": best['lib_description'],
            "lib_spec": best['lib_spec'],
            "lib_brand": best['lib_brand'],
            "lib_model": best['lib_model'],
            "lib_unit": best['lib_unit'],
            "lib_price": lib_price,
            "lib_source": best['lib_source'],
            "difference": diff,
            "diff_percent": diff_pct,
            "status": "manual_check",
            "warning_color": "yellow",
            "match_method": "rapidfuzz_uncertain",
            "match_score": best['score'],
            "candidates": scored_candidates[:5]
        }

    if spec_conflict_candidates:
        best_conf = spec_conflict_candidates[0]
        return {
            "lib_id": best_conf['lib_id'],
            "lib_code": best_conf['lib_code'],
            "lib_description": best_conf['lib_description'],
            "lib_spec": best_conf['lib_spec'],
            "lib_brand": best_conf['lib_brand'],
            "lib_model": best_conf['lib_model'],
            "lib_unit": best_conf['lib_unit'],
            "lib_price": best_conf['lib_price'],
            "lib_source": best_conf['lib_source'],
            "difference": 0.0,
            "diff_percent": 0.0,
            "status": "spec_mismatch",
            "warning_color": "yellow",
            "match_method": "spec_conflict",
            "match_score": best_conf['score'],
            "user_note": best_conf['conflict'],
            "candidates": spec_conflict_candidates[:5]
        }

    # Not found
    return {
        "lib_id": None,
        "lib_code": None,
        "lib_description": None,
        "lib_spec": None,
        "lib_brand": None,
        "lib_model": None,
        "lib_unit": None,
        "lib_price": None,
        "lib_source": None,
        "difference": 0.0,
        "diff_percent": 0.0,
        "status": "not_found",
        "warning_color": "yellow",
        "match_method": "none",
        "match_score": 0.0,
        "candidates": []
    }
