"""Xuất kết quả soát xét VÀO file Excel gốc của NTP.

Nguyên tắc: mở bản gốc lúc nạp (bản sao lưu), giữ nguyên mọi sheet, công thức, định dạng, ô gộp,
dữ liệu; chỉ thêm các cột soát xét bên phải vùng đang dùng của sheet báo giá. Các cột mới dùng công
thức tham chiếu nên Excel tự tính lại khi người dùng sửa."""
import io
import json
import re
import zipfile
from copy import copy
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils import get_column_letter

from backend.library_service import strip_accents
from backend.mep_excel import get_workbook_record
import math

from backend.review_service import STATUS_ADJ, STATUS_DOWN, STATUS_OK, k3_values, round_unit

TITLE = "GIÁ BCHCT SOÁT XÉT"
TITLE3 = "P.KSCP SOÁT XÉT"
# PA3 – Phần 3: các cột soát xét lặp lại lần 2 (P.KSCP)
K3_COLS = [("k3_qty", "KL P.KSCP", 12), ("k3_vt", "Đơn giá VT P.KSCP", 15), ("k3_nc", "Đơn giá NC P.KSCP", 15),
           ("k3_dep", "% KH VT P.KSCP", 10), ("k3_total", "Đơn giá tổng P.KSCP", 16), ("k3_amount", "Thành tiền P.KSCP", 18),
           ("k3_diff", "Chênh lệch (P.KSCP − NTP)", 18), ("k3_pct", "% chênh lệch P.KSCP", 11), ("k3_status", "Kết quả P.KSCP", 12),
           ("k3_note", "Ghi chú P.KSCP", 32)]
# Thứ tự giống Phần 2 trên web: KL | ĐG VT | ĐG NC | % KH | ĐG tổng | Thành tiền | Chênh lệch | Kết quả | Ghi chú
COLS = [("qty", "KL soát xét", 12), ("vt", "Đơn giá VT soát xét", 15), ("nc", "Đơn giá NC soát xét", 15),
        ("dep", "% KH VT soát xét", 10), ("total", "Đơn giá tổng soát xét", 16), ("amount", "Thành tiền soát xét", 18),
        ("diff", "Chênh lệch (soát xét − NTP)", 18), ("pct", "% chênh lệch", 11), ("status", "Kết quả", 12), ("note", "Ghi chú soát xét", 32)]
FILLS = {STATUS_OK: "E7F6EC", STATUS_DOWN: "FDE8E8", STATUS_ADJ: "E0ECFF"}
FONTS = {STATUS_OK: "15803D", STATUS_DOWN: "DC2626", STATUS_ADJ: "1D4ED8"}
THIN = Side(style="thin", color="9CA3AF")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GRAY = Side(style="thin", color="808080")  # khung bảng khi xuất: nét liền, xám
AL_C, AL_R, AL_L = (Alignment(horizontal=h, vertical="center") for h in ("center", "right", "left"))
AL_NOTE = Alignment(horizontal="left", vertical="center", wrap_text=True)
_COLL = {"font": ("_fonts", "fontId"), "fill": ("_fills", "fillId"), "border": ("_borders", "borderId"), "alignment": ("_alignments", "alignmentId")}


class _Styles:
    """Gán kiểu cho ô qua chỉ số đã cache – tránh openpyxl băm lại đối tượng kiểu cho từng ô (chậm)."""
    def __init__(self, wb):
        self.wb, self.cache = wb, {}

    def set(self, cell, kind, obj):
        k = (kind, id(obj))
        hit = self.cache.get(k)
        if hit is None:
            coll, _ = _COLL[kind]
            hit = (getattr(self.wb, coll).add(obj), obj)
            self.cache[k] = hit
        if cell._style is None:
            cell._style = StyleArray()
        setattr(cell._style, _COLL[kind][1], hit[0])
PCT_FMT = '"▲ "0.0%;"▼ "0.0%;0%'  # % chênh lệch: tam giác tăng / giảm

# Tên cột dùng trong công thức cột tự tạo (khớp với giao diện web)
REF_STD = {"KL": "qty", "ĐG VT": "vt", "ĐG NC": "nc", "%KH": "dep", "ĐG tổng": "total", "Thành tiền": "amount",
           "Chênh lệch": "diff", "% Chênh lệch": "pct", "Giá TV": "lib", "KL NTP": "n_qty", "ĐG VT NTP": "n_vt", "ĐG NC NTP": "n_nc",
           "ĐG NTP": "n_price", "TT NTP": "n_amt"}


def _n(v: Any) -> str:
    """Số trong công thức Excel: đủ chữ số, không ký hiệu mũ (1e+06)."""
    t = f"{float(v):.10f}".rstrip("0").rstrip(".")
    return t if t not in ("", "-0") else "0"


def _ref_key(s: str) -> str:
    return re.sub(r"\s+", "", strip_accents(s))


def excel_formula(src: str, refs: Dict[str, str]) -> Optional[str]:
    """Đổi công thức web (=[ĐG VT]*[KL]*1,1; SUM(..;..)) sang công thức Excel với địa chỉ ô của dòng."""
    body = re.sub(r"^\s*=", "", src or "")
    protected = re.sub(r"\[[^\]]*\]", lambda m: m.group(0).replace(",", "\x01"), body)
    if ";" in protected or "(" not in protected:
        protected = re.sub(r"(\d),(\d)", r"\1.\2", protected)
    protected = re.sub(r"\b\d{1,3}(?:\.\d{3})+\b(?!\.)", lambda m: m.group(0).replace(".", ""), protected)
    protected = protected.replace(";", ",").replace("\x01", ",")
    missing = []

    def sub(m):
        k = _ref_key(m.group(1))
        if k not in refs:
            missing.append(m.group(1))
            return "0"
        return refs[k]
    out = re.sub(r"\[([^\]]*)\]", sub, protected)
    return None if missing else "=" + out


def _source_path(rec: Dict[str, Any]) -> Path:
    for key in ("backup_path", "file_path"):  # ưu tiên bản gốc lúc nạp, chưa bị công cụ khác ghi đè
        p = rec.get(key)
        if p and Path(p).exists():
            return Path(p)
    raise FileNotFoundError("Không còn file báo giá gốc trên máy. Nạp lại file NTP rồi soát xét.")


def has_drawings(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as z:
            return any(n.startswith(("xl/drawings/", "xl/charts/", "xl/media/")) for n in z.namelist())
    except zipfile.BadZipFile:
        return False


def _last_used_col(ws, rows: List[int]) -> int:
    last = 0
    for r in rows:
        for c in range(min(ws.max_column or 1, 200), 0, -1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                last = max(last, c)
                break
    for mr in ws.merged_cells.ranges:
        last = max(last, mr.max_col) if mr.min_row <= max(rows) else last
    return last or (ws.max_column or 1)


def _style_from(dst, src, number_format: Optional[str] = None):
    if src is not None and src.has_style:
        # Chép mảng chỉ số kiểu (nhanh gấp nhiều lần chép từng font/viền/màu)
        dst._style = copy(src._style)
    if number_format:
        dst.number_format = number_format


_sc_cache: Dict[int, "_Styles"] = {}


def export_review(package: Dict[str, Any], items: List[Dict[str, Any]], configs: List[Dict[str, Any]]) -> Dict[str, Any]:
    rec = get_workbook_record(package["workbook_id"])
    src = _source_path(rec)
    keep_vba = src.suffix.lower() == ".xlsm"
    wb = openpyxl.load_workbook(src, keep_vba=keep_vba)  # data_only=False: giữ nguyên công thức
    if not configs:
        raise ValueError("Gói chưa có cấu hình sheet soát xét")
    try:
        custom = json.loads(package.get("custom_cols_json") or "[]")
    except ValueError:
        custom = []
    starts = []
    for cfg in configs:
        if cfg["sheet"] not in wb.sheetnames:
            raise ValueError(f"File gốc không còn sheet '{cfg['sheet']}'")
        sheet_items = [i for i in items if i.get("sheet_name") == cfg["sheet"] and not i.get("is_added")]
        added = [i for i in items if i.get("sheet_name") == cfg["sheet"] and i.get("is_added")]
        starts.append(f"{cfg['sheet']}: từ cột {_apply_sheet(wb[cfg['sheet']], cfg, sheet_items, custom, package.get("p2_order"), added, float(package.get("round_unit") or 0), package.get("pa") or "PA2")}")
    _sc_cache.pop(id(wb), None)
    wb.calculation.fullCalcOnLoad = True  # Excel tính lại mọi công thức khi mở
    buf = io.BytesIO()
    wb.save(buf)
    stem = Path(rec.get("original_filename") or "bao-gia.xlsx").stem
    ext = ".xlsm" if keep_vba else ".xlsx"
    return {"data": buf.getvalue(), "filename": f"{stem} - soat xet {package['code']}{ext}",
            "media_type": "application/vnd.ms-excel.sheet.macroEnabled.12" if keep_vba else
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "start_column": "; ".join(starts), "drawings_warning": has_drawings(src)}


def _apply_sheet(ws, cfg: Dict[str, Any], items: List[Dict[str, Any]], custom: Optional[List[Dict[str, Any]]] = None,
                 order: Optional[List[str]] = None, added: Optional[List[Dict[str, Any]]] = None, ru: float = 0, pa: str = "PA2") -> str:
    """Thêm khối cột soát xét bên phải vùng đang dùng của một sheet. Trả chữ cột bắt đầu."""
    mapping = {k: v for k, v in (cfg.get("mapping") or {}).items() if v}
    header_row = int(cfg.get("header_row") or 1)
    item_rows = [int(i["excel_row"]) for i in items]
    scan = list(range(1, (max(item_rows) if item_rows else header_row) + 1))
    # chừa 1 cột trống tách phần NTP và phần soát xét; dùng đúng cột web đã hiển thị (p2_start) để công thức ô khớp địa chỉ
    start = int(cfg.get("p2_start") or 0) or (_last_used_col(ws, scan) + 2)
    digits = -int(round(math.log10(ru))) if ru else None
    rnd = (lambda expr: f"ROUND({expr},{digits})") if ru else (lambda expr: expr)
    cols = list(COLS)
    for c in custom or []:  # cột người dùng tạo thêm
        cols.append(("x_" + c["id"], c["label"], 16 if c.get("type") == "number" else 24))
    if order:  # theo thứ tự cột Phần 2 người dùng sắp xếp
        alias = {"amt": "amount"}
        pos = {alias.get(k, k): i for i, k in enumerate(order)}
        cols.sort(key=lambda x: pos.get(x[0], 999))
    types = {"x_" + c["id"]: c.get("type") for c in custom or []}
    n2 = len(cols)
    if pa == "PA3":
        cols += K3_COLS
    col = {key: start + i for i, (key, _, _) in enumerate(cols)}
    L = {k: get_column_letter(v) for k, v in col.items()}

    # Tiêu đề: chép kiểu ô tiêu đề cột Hạng mục của file gốc
    head_src = ws.cell(row=header_row, column=int(mapping.get("description") or 1))
    for key, label, width in cols:
        c = ws.cell(row=header_row, column=col[key], value=str(label).upper())  # tiêu đề chữ hoa
        _style_from(c, head_src)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.font = Font(bold=True, color="1E3A8A", name=head_src.font.name if head_src.has_style else None,
                      size=head_src.font.size if head_src.has_style else None)
        c.fill = PatternFill("solid", start_color="EAF2FF")
        c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        ws.column_dimensions[L[key]].width = width
    if header_row > 1:
        top = header_row - 1
        free = all(ws.cell(row=top, column=col[k]).value in (None, "") for k, _, _ in cols)
        if free:
            bands = [(cols[:n2], TITLE, "1E3A8A")] + ([(cols[n2:], TITLE3, "9A3412")] if len(cols) > n2 else [])
            for part, title, color in bands:
                c1, c2 = col[part[0][0]], col[part[-1][0]]
                ws.merge_cells(start_row=top, start_column=c1, end_row=top, end_column=c2)
                t = ws.cell(row=top, column=c1, value=title)
                t.font = Font(bold=True, color="FFFFFF")
                t.fill = PatternFill("solid", start_color=color)
                t.alignment = Alignment(horizontal="center", vertical="center")

    def ref(field: str, r: int) -> Optional[str]:
        c = mapping.get(field)
        return f"{get_column_letter(int(c))}{r}" if c else None

    for it in items:
        r = int(it["excel_row"])
        price_src = ws.cell(row=r, column=int(mapping.get("price") or mapping.get("material_price") or col["qty"]))
        qty_src = ws.cell(row=r, column=int(mapping.get("quantity") or col["qty"]))
        text_src = ws.cell(row=r, column=int(mapping.get("description") or 1))
        money_fmt = price_src.number_format if price_src.number_format != "General" else "#,##0"
        qty_fmt = qty_src.number_format if qty_src.number_format != "General" else "#,##0.###"
        has_parts = any(it.get(k) is not None for k in ("rv_material", "rv_labor", "rv_machine"))
        vals = {
            "qty": it.get("rv_quantity") if it.get("rv_quantity") is not None else it.get("boq_quantity"),
            "vt": round_unit(it.get("rv_material"), ru), "nc": round_unit(it.get("rv_labor"), ru), "total": it.get("rv_unit_total"),
        }
        dep = it.get("rv_dep_pct") if it.get("rv_dep_pct") is not None else it.get("boq_dep_pct")
        vals["dep"] = dep
        if has_parts and it.get("rv_total") is None:
            # Đơn giá = VT × %KH + NC (khấu hao/máy của file cũ, nếu có, cộng thêm dạng số)
            vt = f"{L['vt']}{r}*{L['dep']}{r}" if dep is not None else f"{L['vt']}{r}"
            mac = it.get("rv_machine")
            vals["total"] = "=" + rnd(f"{vt}+{L['nc']}{r}" + (f"+{_n(round_unit(mac, ru))}" if mac else ""))
        # Đơn giá KSCP mặc định là công thức ROUND(…,-2): giá lấy theo NTP → tham chiếu ô NTP (=ROUND(E7,-2)),
        # giá thư viện / nhập tay → =ROUND(số,-2)
        if ru:
            def _rp(v, ntp_v, field):
                if v is None:
                    return None
                src = ref(field, r) if ntp_v is not None and abs(float(v) - float(ntp_v)) < 1e-6 else None
                return "=" + rnd(src or f"{_n(v)}")
            # KL soát xét: =ROUND(D7,0) khi theo NTP, =ROUND(số,0) khi sửa
            qv = it.get("rv_quantity") if it.get("rv_quantity") is not None else it.get("boq_quantity")
            if qv is not None:
                qsrc = ref("quantity", r) if it.get("boq_quantity") is not None and abs(float(qv) - float(it["boq_quantity"])) < 1e-9 else None
                vals["qty"] = f"=ROUND({qsrc or f'{_n(qv)}'},0)"
            vals["vt"] = _rp(it.get("rv_material"), it.get("boq_material"), "material_price")
            vals["nc"] = _rp(it.get("rv_labor"), it.get("boq_labor"), "labor_price")
            if not has_parts or it.get("rv_total") is not None:
                tv = it.get("rv_total") if it.get("rv_total") is not None else it.get("boq_price")
                vals["total"] = _rp(tv, it.get("boq_price"), "price")
        vals |= {
            "amount": f"={L['total']}{r}*{L['qty']}{r}",  # Thành tiền = Đơn giá × Khối lượng
        }
        ntp_amount = ref("amount", r)
        if not ntp_amount and ref("quantity", r) and ref("price", r):
            ntp_amount = f"{ref('quantity', r)}*{ref('price', r)}"
        vals["diff"] = f"={L['amount']}{r}-{ntp_amount}" if ntp_amount else it.get("diff_amount")
        vals["pct"] = f'=IF(N({ntp_amount})=0,"",{L["diff"]}{r}/{ntp_amount})' if ntp_amount else (
            it["diff_amount"] / it["ntp_amount"] if it.get("ntp_amount") else None)
        vals["status"] = it.get("review_status")
        vals["note"] = it.get("rv_note") or None
        try:
            extra = json.loads(it.get("rv_extra") or "{}")
        except ValueError:
            extra = {}
        # Tham chiếu ô cho công thức cột tự tạo
        refs: Dict[str, str] = {}
        own = {"qty": "qty", "vt": "vt", "nc": "nc", "total": "total", "amount": "amount", "diff": "diff", "pct": "pct"}
        for name, key in REF_STD.items():
            if key in own:
                refs[_ref_key(name)] = f"{L[own[key]]}{r}"
            elif key == "dep":
                refs[_ref_key(name)] = f"{L['dep']}{r}" if dep is not None else "1"
            elif key == "lib":
                refs[_ref_key(name)] = f"{_n(it.get('lib_price') or 0)}"
            else:
                field = {"n_qty": "quantity", "n_vt": "material_price", "n_nc": "labor_price", "n_price": "price", "n_amt": "amount"}[key]
                lit = {"n_qty": it.get("boq_quantity"), "n_vt": it.get("boq_material"), "n_nc": it.get("boq_labor"),
                       "n_price": it.get("boq_price"), "n_amt": float(it.get("boq_quantity") or 0) * float(it.get("boq_price") or 0)}[key]
                refs[_ref_key(name)] = ref(field, r) or f"{_n(lit or 0)}"
        for c in custom or []:
            if c.get("type") != "text":
                refs[_ref_key(c["label"])] = f"{L['x_' + c['id']]}{r}"
        for k in types:
            if types[k] == "formula":
                f = next((c.get("formula") for c in custom or [] if "x_" + c["id"] == k), "")
                vals[k] = excel_formula(f, refs)
            else:
                vals[k] = extra.get(k[2:])
        # Số NTP đã sửa trên bảng (Phần 1) → ghi vào đúng ô gốc (giá trị hoặc công thức); đơn giá NTP là công thức gốc
        # (vd =E7*G7+F7) mà chỉ sửa VT/NC/%KH thì giữ công thức để Excel tự tính
        ntp_map = {"boq_quantity": "quantity", "boq_material": "material_price", "boq_labor": "labor_price",
                   "boq_dep_pct": "depreciation_pct", "boq_price": "price", "boq_dep_done": "dep_done_pct"}
        # STT / Mô tả công việc sửa trên bảng → ghi vào ô gốc
        for tk, mk in (("t_stt", "stt"), ("t_desc", "description")):
            if extra.get(tk) and mapping.get(mk):
                ws.cell(row=r, column=int(mapping[mk])).value = extra[tk]
        # Ngày bắt đầu / kết thúc sửa trên bảng → ghi vào ô ngày gốc (kiểu ngày dd/mm/yyyy)
        for dk, mk in (("d_start", "start_date"), ("d_end", "end_date")):
            if extra.get(dk) and mapping.get(mk):
                from backend.depreciation import parse_date
                dd = parse_date(extra[dk])
                if dd:
                    c0 = ws.cell(row=r, column=int(mapping[mk]))
                    c0.value = dd
                    c0.number_format = "dd/mm/yyyy"
        if it.get("boq_dep_done"):  # có % ĐÃ KH → ĐG NC soát xét để trống
            vals["nc"] = None
        for f, mk in ntp_map.items():
            if "o_" + f not in extra or not mapping.get(mk):
                continue
            cell = ws.cell(row=r, column=int(mapping[mk]))
            fm = extra.get("f_" + f)
            if f == "boq_price" and not fm and isinstance(cell.value, str) and cell.value.startswith("=") \
                    and any("o_" + k in extra for k in ("boq_material", "boq_labor", "boq_dep_pct")):
                continue
            ex = excel_formula(fm, refs) if isinstance(fm, str) and fm.startswith("=") else None
            cell.value = ex or it.get(f)
        if pa == "PA1":  # Dự toán ĐNTC: Đơn giá (K) = E × %KH + F, Thành tiền (L) = K × KL – công thức Excel nếu ô gốc chưa có công thức
            e_, f_, h_, d_ = ref("material_price", r), ref("labor_price", r), ref("depreciation_pct", r), ref("quantity", r)
            for mk, fx in (("price", (f"={e_}*IF(N({h_})=0,1,{h_})+N({f_})" if h_ else f"={e_}+N({f_})") if e_ and f_ else None),
                           ("amount", f"={ref('price', r)}*{d_}" if ref("price", r) and d_ else None)):
                if not fx or not mapping.get(mk) or extra.get("f_boq_price" if mk == "price" else "f_x_ov_ntp_amt") \
                        or (mk == "amount" and isinstance(extra.get("ov_ntp_amt"), (int, float))):
                    continue
                cur0 = ws.cell(row=r, column=int(mapping[mk])).value
                if not (isinstance(cur0, str) and cur0.startswith("=")):
                    ws.cell(row=r, column=int(mapping[mk])).value = fx
        # Công thức gõ trên bảng (=M7*1,1, =SUM(E7:E9)…) → công thức Excel thật; đơn giá bọc ROUND theo làm tròn
        fmap = {"f_rv_quantity": ("qty", False), "f_rv_material": ("vt", True), "f_rv_labor": ("nc", True),
                "f_rv_dep_pct": ("dep", False), "f_rv_total": ("total", True)}
        for fk, fv in extra.items():
            if not (fk.startswith("f_") and isinstance(fv, str) and fv.startswith("=")):
                continue
            key, is_price = fmap.get(fk, (fk[2:], False)) if not fk.startswith("f_x_") else (fk[2:], False)
            if key not in col or (key == "total" and has_parts and it.get("rv_total") is None):
                continue
            ex = excel_formula(fv, refs)
            if ex:
                vals[key] = "=" + rnd(ex[1:]) if is_price else (f"=ROUND({ex[1:]},0)" if key == "qty" and ru else ex)
        # Thành tiền gõ / công thức riêng: soát xét → ô Thành tiền Phần 2; NTP → ô Thành tiền gốc của file
        for ok, fk in (("ov_rv_amt", "f_x_ov_rv_amt"), ("ov_ntp_amt", "f_x_ov_ntp_amt")):
            if not isinstance(extra.get(ok), (int, float)):
                continue
            fm = extra.get(fk)
            ex = excel_formula(fm, refs) if isinstance(fm, str) and fm.startswith("=") else None
            if ok == "ov_rv_amt":
                vals["amount"] = ex or extra[ok]
            elif mapping.get("amount"):
                ws.cell(row=r, column=int(mapping["amount"])).value = ex or extra[ok]
        if pa == "PA3":  # Phần 3: ô chưa sửa = tham chiếu ô Phần 2 cùng dòng (theo BCHCT), ô đã sửa = số / công thức gõ
            kv = k3_values(it)
            def _k3(name, base, is_price=False):
                fv = extra.get("f_x_k3_" + name)
                ex = excel_formula(fv, refs) if isinstance(fv, str) and fv.startswith("=") else None
                if ex:
                    return "=" + rnd(ex[1:]) if is_price else ex
                v = extra.get("k3_" + name)
                if isinstance(v, (int, float)):
                    return "=" + rnd(f"{_n(v)}") if is_price and ru else v
                return f"={L[base]}{r}" if base else None
            vals["k3_qty"] = _k3("qty", "qty")
            vals["k3_vt"] = _k3("vt", "vt", True)
            vals["k3_nc"] = None if it.get("boq_dep_done") else _k3("nc", "nc", True)
            vals["k3_dep"] = _k3("dep", "dep")
            tv = _k3("total", None, True)
            if tv is None:
                parts = any(v is not None for v in (kv["vt"], kv["nc"], it.get("rv_machine")))
                touched = any(isinstance(extra.get(k), (int, float)) or extra.get("f_x_" + k) for k in ("k3_vt", "k3_nc", "k3_dep"))
                if parts and (it.get("rv_total") is None or touched):
                    vt3 = f"{L['k3_vt']}{r}*{L['k3_dep']}{r}" if kv["dep"] is not None else f"{L['k3_vt']}{r}"
                    mac = it.get("rv_machine")
                    tv = "=" + rnd(f"{vt3}+{L['k3_nc']}{r}" + (f"+{_n(round_unit(mac, ru))}" if mac else ""))
                else:
                    tv = f"={L['total']}{r}"
            vals["k3_total"] = tv
            av = _k3("amt", None)
            vals["k3_amount"] = av if av is not None else f"={L['k3_total']}{r}*{L['k3_qty']}{r}"
            vals["k3_diff"] = f"={L['k3_amount']}{r}-{ntp_amount}" if ntp_amount else kv["diff"]
            vals["k3_pct"] = f'=IF(N({ntp_amount})=0,"",{L["k3_diff"]}{r}/{ntp_amount})' if ntp_amount else None
            vals["k3_status"] = kv["status"]
            vals["k3_note"] = kv["note"] or None
        st_cells = extra.get("st") if isinstance(extra.get("st"), dict) else {}
        sc = _sc_cache.setdefault(id(ws.parent), _Styles(ws.parent))
        for key0, _, _ in cols:
            c = ws.cell(row=r, column=col[key0], value=vals[key0])
            key = key0[3:] if key0.startswith("k3_") else key0
            if key == "dep":
                _style_from(c, price_src, "0%")
                sc.set(c, "alignment", AL_C)
            elif key == "pct":
                _style_from(c, price_src, PCT_FMT)
                sc.set(c, "alignment", AL_R)
                v = it["diff_amount"] / it["ntp_amount"] if it.get("ntp_amount") else 0
                if abs(v) >= 0.0005:
                    c.font = Font(color="DC2626" if v < 0 else "1D4ED8", name=c.font.name, size=c.font.size)
            elif key in ("status", "note") or (key in types and types[key] == "text"):
                _style_from(c, text_src)
                sc.set(c, "alignment", AL_C if key == "status" else AL_NOTE if key == "note" else AL_L)
            else:
                _style_from(c, price_src if key != "qty" else qty_src, qty_fmt if key == "qty" else money_fmt)
                sc.set(c, "alignment", AL_R)
            sc.set(c, "border", BORDER_ALL)
        # Định dạng ô người dùng đặt trên bảng (thanh công cụ): áp vào ô Excel tương ứng
        for cid, sty in st_cells.items():
            if cid.startswith("p1:c") and cid[4:].isdigit():
                tc = ws.cell(row=r, column=int(cid[4:]))
            elif cid.startswith("p2:"):
                k2 = {"amt": "amount", "k3_amt": "k3_amount"}.get(cid[3:], cid[3:])
                if k2 not in col:
                    continue
                tc = ws.cell(row=r, column=col[k2])
            else:
                continue
            f0 = tc.font
            tc.font = Font(name=f0.name, size=f0.size, bold=bool(sty.get("b")) or f0.bold, italic=bool(sty.get("i")) or f0.italic,
                           underline="single" if sty.get("u") else f0.underline,
                           color=sty["c"][1:] if sty.get("c") else f0.color)
            if sty.get("bg"):
                tc.fill = PatternFill("solid", start_color=sty["bg"][1:])
            if sty.get("al"):
                tc.alignment = Alignment(horizontal={"l": "left", "c": "center", "r": "right"}[sty["al"]], vertical="center",
                                         wrap_text=tc.alignment.wrap_text)
            if sty.get("nf"):
                tc.number_format = {"int": "#,##0", "d2": "#,##0.00", "pct": "0.0%"}[sty["nf"]]
        for sk, st in (("status", it.get("review_status")), ("k3_status", vals.get("k3_status"))):
            if sk in col and st in FILLS:
                s = ws.cell(row=r, column=col[sk])
                s.fill = PatternFill("solid", start_color=FILLS[st])
                s.font = Font(bold=True, color=FONTS[st])

    total_row = None
    if item_rows:
        first, last = min(item_rows), max(item_rows)
        # Dòng tổng của file gốc (vd "TỔNG GIÁ TRỊ (Chưa VAT)" có thành tiền) → ghi tổng soát xét CÙNG HÀNG, tô xanh nhạt
        tr = None
        amt_col = int(mapping["amount"]) if mapping.get("amount") else None
        c_from, c_to = int(cfg.get("first_col") or 1), int(cfg.get("last_col") or (min(col.values()) - 2))
        for rr in range(last + 1, (ws.max_row or last) + 1):
            texts = " ".join(strip_accents(str(ws.cell(row=rr, column=cc).value)) for cc in range(c_from, c_to + 1)
                             if isinstance(ws.cell(row=rr, column=cc).value, str))
            has_amt = amt_col is None or ws.cell(row=rr, column=amt_col).value not in (None, "")
            if re.search(r"\btong\b", texts) and has_amt:
                tr = rr
                break
        same_row = tr is not None
        total_row = tr if tr is not None else max(ws.max_row, last) + 1
        if tr is None:
            tr = max(ws.max_row, last) + 1
        lab = ws.cell(row=tr, column=col["total"], value="Tổng soát xét")
        lab.font = Font(bold=True)
        if same_row:
            blue = PatternFill("solid", start_color="DDEBF7")
            for cc in range(min(col.values()), max(col.values()) + 1):
                ws.cell(row=tr, column=cc).fill = blue
        for key in ("amount", "diff", *[k for k, t in types.items() if t != "text"], *(("k3_amount", "k3_diff") if pa == "PA3" else ())):
            c = ws.cell(row=tr, column=col[key], value=f"=SUBTOTAL(9,{L[key]}{first}:{L[key]}{last})")
            c.font = Font(bold=True)
            c.number_format = "#,##0"
            c.border = Border(top=Side(style="double", color="1E3A8A"))
        A, D = f"{L['amount']}{tr}", f"{L['diff']}{tr}"
        for pk, ak, dk in (("pct", "amount", "diff"), *((("k3_pct", "k3_amount", "k3_diff"),) if pa == "PA3" else ())):
            A, D = f"{L[ak]}{tr}", f"{L[dk]}{tr}"
            c = ws.cell(row=tr, column=col[pk], value=f'=IF(({A}-{D})=0,"",{D}/({A}-{D}))')
            c.font = Font(bold=True)
            c.number_format = PCT_FMT
            c.border = Border(top=Side(style="double", color="1E3A8A"))

    if added:  # dòng KSCP thêm vào: ghi thành khối riêng dưới bảng, không chèn vào giữa để giữ nguyên công thức file gốc
        r0 = ws.max_row + 2
        t = ws.cell(row=r0, column=int(mapping.get("description") or min(col.values())), value="CÁC DÒNG KSCP BỔ SUNG")
        t.font = Font(bold=True, color="1E3A8A")
        for n, it in enumerate(added, start=1):
            r = r0 + n
            for field, v in (("stt", it.get("boq_stt")), ("description", it.get("boq_description")), ("unit", it.get("boq_unit"))):
                if mapping.get(field):
                    ws.cell(row=r, column=int(mapping[field]), value=v)
            try:
                extra = json.loads(it.get("rv_extra") or "{}")
            except ValueError:
                extra = {}
            dep = it.get("rv_dep_pct")
            vals = {"qty": it.get("rv_quantity"), "vt": round_unit(it.get("rv_material"), ru), "nc": round_unit(it.get("rv_labor"), ru), "dep": dep,
                    "total": it.get("rv_unit_total"), "amount": it.get("rv_amount"), "diff": it.get("diff_amount"),
                    "status": "Bổ sung", "note": it.get("rv_note")}
            for k in types:
                vals[k] = extra.get(k[2:]) if types[k] != "formula" else None
            for key, _, _ in cols:
                c = ws.cell(row=r, column=col[key], value=vals.get(key))
                c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
                if key == "dep":
                    c.number_format = "0%"
                elif key not in ("status", "note") and types.get(key) != "text":
                    c.number_format = "#,##0"
    # Phông khối soát xét: Times New Roman – tiêu đề cỡ 13, nội dung cỡ 11 (giữ đậm / nghiêng / màu đã có)
    c_lo, c_hi = min(col.values()), max(col.values())
    scf = _sc_cache.setdefault(id(ws.parent), _Styles(ws.parent))
    font_cache: Dict[tuple, Font] = {}
    for rr in range(max(1, header_row - 1), (ws.max_row or header_row) + 1):
        size = 13 if rr <= header_row or rr == total_row else 11
        for cc in range(c_lo, c_hi + 1):
            cell = ws.cell(row=rr, column=cc)
            f0 = cell.font
            color = f0.color.rgb if f0 is not None and f0.color is not None and isinstance(f0.color.rgb, str) else None
            key = (size, bool(f0.bold), bool(f0.italic), f0.underline, color)
            fnt = font_cache.get(key)
            if fnt is None:
                fnt = font_cache[key] = Font(name="Times New Roman", size=size, bold=key[1], italic=key[2], underline=key[3], color=color)
            scf.set(cell, "font", fnt)
    # Dòng STT mẹ (1, 1.1, 1.1.1… không phải hạng mục): Thành tiền / Chênh lệch soát xét = SUBTOTAL(9, các dòng con) như file NTP
    if item_rows and mapping.get("stt"):
        normS = lambda v: str(v or "").strip().replace(",", ".").rstrip(".")
        item_stt = [(int(i["excel_row"]), normS(i.get("boq_stt"))) for i in items if normS(i.get("boq_stt"))]
        item_set2 = set(item_rows)
        sum_keys = ["amount", "diff", *[k for k, t in types.items() if t != "text"], *(("k3_amount", "k3_diff") if pa == "PA3" else ())]
        amt_c = int(mapping["amount"]) if mapping.get("amount") else None
        for rr in range(header_row + 1, (total_row or max(item_rows) + 1)):
            if rr in item_set2:  # hạng mục có dòng con mà thành tiền file là SUBTOTAL → cũng là dòng mẹ
                fv = ws.cell(row=rr, column=amt_c).value if amt_c else None
                if not (isinstance(fv, str) and "SUBTOTAL" in fv.upper()):
                    continue
            st0 = normS(ws.cell(row=rr, column=int(mapping["stt"])).value)
            if not st0:
                continue
            kids = [x for x, sv in item_stt if sv.startswith(st0 + ".") and x != rr]
            if not kids:
                continue
            a, b2 = min(kids), max(kids)
            for key in sum_keys:
                c = ws.cell(row=rr, column=col[key], value=f"=SUBTOTAL(9,{L[key]}{a}:{L[key]}{b2})")
                c.number_format = "#,##0"
                c.font = Font(bold=True)
                c.alignment = AL_R
            ntp_ref = f"{get_column_letter(int(mapping['amount']))}{rr}" if mapping.get("amount") else f"({L['amount']}{rr}-{L['diff']}{rr})"
            for pk, dk in (("pct", "diff"), *((("k3_pct", "k3_diff"),) if pa == "PA3" else ())):
                c = ws.cell(row=rr, column=col[pk], value=f'=IF(N({ntp_ref})=0,"",{L[dk]}{rr}/{ntp_ref})')
                c.number_format = PCT_FMT
                c.font = Font(bold=True)
                c.alignment = AL_R
    # Dòng hạng mục mẹ (dòng nhóm: A, 1, 1.1, 1.1.1…) → tô xám nhạt phần soát xét cho dễ nhìn
    if item_rows:
        grey = PatternFill("solid", start_color="E9EBEF")
        item_set = set(item_rows)
        p_lo, p_hi = int(cfg.get("first_col") or 1), int(cfg.get("last_col") or (c_lo - 2))
        end = (total_row - 1) if total_row else max(item_rows)
        # Hạng mục có STT là "mẹ" của dòng khác (vd 1.2.1 có 1.2.1.1…) cũng tô xám
        norm = lambda v: str(v or "").strip().replace(",", ".").rstrip(".")
        stts = [norm(i.get("boq_stt")) for i in items if norm(i.get("boq_stt"))]
        parent_rows = {int(i["excel_row"]) for i in items
                       if norm(i.get("boq_stt")) and any(x.startswith(norm(i.get("boq_stt")) + ".") for x in stts)}
        for rr in range(header_row + 1, end + 1):
            if rr in item_set and rr not in parent_rows:
                continue
            if all(ws.cell(row=rr, column=cc).value in (None, "") for cc in range(p_lo, p_hi + 1)):
                continue
            # Giữ định dạng của file: dòng mẹ có màu nền trong file NTP → phần soát xét cùng màu; không có thì xám nhạt
            src = ws.cell(row=rr, column=int(mapping.get("description") or p_lo))
            sf = src.fill if src.has_style and src.fill is not None and src.fill.fill_type == "solid" else None
            src_rgb = sf.start_color.rgb if sf is not None and isinstance(sf.start_color.rgb, str) else None
            # màu theme / indexed (không phải mã RGB) cũng giữ nguyên
            is_white = sf is not None and src_rgb is None and getattr(sf.start_color, "theme", None) == 0 and not (sf.start_color.tint or 0)
            use = copy(sf) if sf is not None and not is_white and (src_rgb is None or src_rgb not in ("00000000", "FFFFFFFF")) else grey
            for cc in range(c_lo, c_hi + 1):
                ws.cell(row=rr, column=cc).fill = use
    # Phần 1 (bảng NTP): dòng tiêu đề và dòng tổng cũng Times New Roman cỡ 13 cho đồng bộ 2 phần
    p1_lo, p1_hi = int(cfg.get("first_col") or 1), int(cfg.get("last_col") or (c_lo - 2))
    for rr in [r for r in range(max(1, header_row - 1), header_row + 1)] + ([total_row] if total_row else []):
        for cc in range(p1_lo, p1_hi + 1):
            cell = ws.cell(row=rr, column=cc)
            if cell.value in (None, "") and rr != total_row:
                continue
            f0 = cell.font
            color = f0.color.rgb if f0 is not None and f0.color is not None and isinstance(f0.color.rgb, str) else None
            key = (13, bool(f0.bold), bool(f0.italic), f0.underline, color)
            fnt = font_cache.get(key) or font_cache.setdefault(key, Font(name="Times New Roman", size=13, bold=key[1], italic=key[2], underline=key[3], color=color))
            scf.set(cell, "font", fnt)
    # Tự căn: độ rộng cột soát xét theo nội dung (kết quả số ước tính theo giá trị đã tính) + trang in vừa khổ ngang
    widths: Dict[int, float] = {}
    for key, label, w0 in cols:
        cc = col[key]
        longest_word = max((len(wd) for wd in str(label).split()), default=6)
        widths[cc] = max(float(w0), longest_word * 1.25 + 2)
    sample_val = {"qty": "rv_quantity", "vt": "rv_material", "nc": "rv_labor", "total": "rv_unit_total", "amount": "rv_amount", "diff": "diff_amount"}
    for it in items:
        for key, f in sample_val.items():
            v = it.get(f)
            if isinstance(v, (int, float)) and key in col:
                widths[col[key]] = max(widths[col[key]], len(f"{abs(v):,.0f}") * 1.15 + 3)
    for key, _, _ in cols:  # Phần 3 rộng như cột tương ứng ở Phần 2
        if key.startswith("k3_") and key[3:] in col:
            widths[col[key]] = max(widths.get(col[key], 0), widths.get(col[key[3:]], 0))
    for cc, wv in widths.items():
        ws.column_dimensions[get_column_letter(cc)].width = min(wv, 40)
    try:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A3 if (c_hi - (int(cfg.get("first_col") or 1))) > 18 else ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = ws.page_margins.right = 0.3
        ws.page_margins.top = ws.page_margins.bottom = 0.5
        ws.print_options.horizontalCentered = True
        ws.print_title_rows = f"{max(1, header_row - 1)}:{header_row}"  # lặp dòng tiêu đề mỗi trang
    except Exception:
        pass
    # Cột ngăn cách phần NTP và phần soát xét: tô xanh nhạt, rộng 2
    sep = min(col.values()) - 1
    if sep >= 1:
        ws.column_dimensions[get_column_letter(sep)].width = 2
        fill = PatternFill("solid", start_color="BDD7EE")
        for rr in range(1, (ws.max_row or 1) + 1):
            cell = ws.cell(row=rr, column=sep)
            if cell.value in (None, ""):
                cell.fill = fill
    # Kẻ khung: mọi ô của bảng (Phần 1 + cột ngăn cách + Phần 2), từ dòng tiêu đề tới dòng cuối → nét liền mảnh màu xám
    grid = Border(left=GRAY, right=GRAY, top=GRAY, bottom=GRAY)
    b_lo = int(cfg.get("first_col") or 1)
    top = max(1, header_row - 1)
    if all(ws.cell(row=top, column=cc).value in (None, "") for cc in range(b_lo, c_hi + 1)) and \
            not any(m.min_row <= top <= m.max_row for m in ws.merged_cells.ranges):
        top = header_row
    end = top
    for rr in range((ws.max_row or header_row), header_row - 1, -1):
        if any(ws.cell(row=rr, column=cc).value not in (None, "") for cc in range(b_lo, c_hi + 1)):
            end = rr
            break
    for rr in range(top, end + 1):
        for cc in range(b_lo, c_hi + 1):
            scf.set(ws.cell(row=rr, column=cc), "border", grid)
    if pa == "PA1":  # Dự toán ĐNTC: chỉ có Phần 1 (giá đã ghi vào đúng ô E, F, H của file) → bỏ khối cột soát xét
        lo = min(col.values())
        for m in list(ws.merged_cells.ranges):
            if m.max_col >= lo:
                ws.unmerge_cells(str(m))
        ws.delete_cols(lo - 1, (ws.max_column or lo) - lo + 2)
        return get_column_letter(lo) + " (Dự toán ĐNTC – chỉ Phần 1)"
    return get_column_letter(min(col.values()))
