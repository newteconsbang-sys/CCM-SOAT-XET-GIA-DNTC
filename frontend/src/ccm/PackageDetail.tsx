import React, { useEffect, useMemo, useRef, useState } from 'react';
import { mepApi, mepDownload } from '../mepApi';
import { evalFormula, checkFormula, refKey, FUNCTION_NAMES, FUNC_INFO, roundUnit } from './formula';
import { Badge, Icon, Pkg, STATUSES, Version, dmy, money, num, parseVn, pct, toast, PAS } from './ui';

type RV = { rv_quantity: number | null; rv_material: number | null; rv_labor: number | null; rv_machine: number | null; rv_total: number | null; rv_note: string | null; rv_dep_pct?: number | null };
type Item = RV & {
  id: number; excel_row: number; boq_stt: string | null; boq_code: string; boq_description: string; boq_unit: string;
  boq_quantity: number; boq_price: number; boq_material: number | null; boq_labor: number | null; boq_machine: number | null;
  boq_amount: number | null; boq_note: string | null; boq_dep_pct?: number | null;
  lib_code: string | null; lib_description: string | null; lib_unit: string | null; lib_price: number | null; lib_material: number | null;
  lib_labor: number | null; lib_machine: number | null; lib_flag: 'none' | 'uncertain' | 'higher' | 'lower' | 'ok'; lib_diff_percent: number | null;
  status: string; match_score: number | null; rv_touched: number; rv_extra?: string | null; st?: Record<string, CellStyle> | null; is_added?: number; add_stt?: string | null; add_desc?: string | null; add_unit?: string | null;
  [k: `x_${string}`]: any;
  [f: `f_${string}`]: any;
};
type SheetCfg = { sheet: string; header_row: number; mapping: Record<string, number | null>; headers: Record<string, string>; first_col: number | null; last_col: number | null; p2_start?: number };
type SRow = { row: number; kind: 'item' | 'group'; cells: any[]; item_id: number | null; sid?: number; pos?: number; f?: (string | null)[] | null };
type CustomCol = { id: string; label: string; type: 'number' | 'text' | 'formula'; formula?: string };
// Tên cột dùng trong công thức: [KL], [ĐG VT]… (cột soát xét) và [KL NTP], [ĐG NTP]… (giá NTP báo)
const REF_STD: [string, string][] = [
  ['KL', 'qty'], ['ĐG VT', 'vt'], ['ĐG NC', 'nc'], ['%KH', 'dep'], ['ĐG tổng', 'total'], ['Thành tiền', 'amt'], ['Chênh lệch', 'diff'], ['% Chênh lệch', 'pct'], ['Giá TV', 'lib'],
  ['KL NTP', 'n_qty'], ['ĐG VT NTP', 'n_vt'], ['ĐG NC NTP', 'n_nc'], ['ĐG NTP', 'n_price'], ['TT NTP', 'n_amt'],
];
type Detail = Pkg & { kh_region?: string; kh_info?: Record<string, { months: number | null; can: boolean; group: string; pct: number | null; sheet?: string | null; t?: number; table?: string | null }>; custom_columns?: CustomCol[]; p2_order?: string[]; round_unit?: number; items: Item[] & { sheet_name?: string }[]; sheets?: string[]; session?: { threshold_percent: number } | null; header_row?: number | null; mapping_json?: string | null;
  sheet_configs?: SheetCfg[]; structure?: Record<string, SRow[]> };
type P1Col = { key: string; label: string; field?: string; width: number; cls: string; get: (r: SRow, it?: Item) => React.ReactNode };
type CellStyle = { b?: boolean; i?: boolean; u?: boolean; c?: string; bg?: string; al?: 'l' | 'c' | 'r'; nf?: 'int' | 'd2' | 'pct' };
type Edit = Partial<RV> & { st?: Record<string, CellStyle> | null; [k: `x_${string}`]: any; [f: `f_${string}`]: any; add_stt?: string | null; add_desc?: string | null; add_unit?: string | null };
type Hist = { id: number; field: string; field_label: string; old_value: string | null; new_value: string | null; item_label: string | null; changed_by: string; changed_at: string };
type MapKey = 'stt' | 'description' | 'unit' | 'quantity' | 'material_price' | 'labor_price' | 'machine_price' | 'depreciation_pct' | 'dep_done_pct' | 'start_date' | 'end_date' | 'price' | 'amount' | 'note' | 'code';
const MAP_FIELDS: [MapKey, string][] = [
  ['stt', 'STT'], ['description', 'Hạng mục *'], ['unit', 'ĐVT'], ['quantity', 'Khối lượng'],
  ['material_price', 'Đơn giá VT'], ['dep_done_pct', '% đã KH'], ['depreciation_pct', '% KH vật tư'], ['start_date', 'Thời gian bắt đầu'], ['end_date', 'Thời gian kết thúc'], ['labor_price', 'Đơn giá NC'], ['machine_price', 'Khấu hao'], ['price', 'Đơn giá tổng'],
  ['amount', 'Thành tiền'], ['note', 'Ghi chú'], ['code', 'Mã hiệu'],
];
// Kết quả: NTP cao hơn soát xét → Cao (đỏ); thấp hơn → Thấp (xanh dương); bằng → Hợp lý (xanh lá)
const STATUS_CLS: Record<string, string> = { 'Hợp lý': 'ok', 'Cao': 'bad', 'Thấp': 'low', 'Bổ sung': 'add' };
const FILTERS: [string, string][] = [
  ['', 'Tất cả'], ['Hợp lý', 'Hợp lý'], ['Cao', 'NTP cao hơn (Cao)'], ['Thấp', 'NTP thấp hơn (Thấp)'],
  ['lib:higher', 'NTP cao hơn thư viện'], ['lib:uncertain', 'Khớp thư viện chưa chắc'], ['lib:none', 'Không có trong thư viện'], ['edited', 'Đã sửa tay'],
];
// Độ rộng cột: 4 cột cố định | Phần 1 (8) | Phần 2 (10). 0 = Hạng mục co giãn.
// Phần 2: KL | ĐG VT | ĐG NC | % KH | ĐG tổng | Thành tiền | Chênh lệch | Giá TV | Kết quả | Ghi chú
const P2_WIDTHS = [70, 100, 100, 64, 104, 112, 100, 96, 86, 180];
// Cột chuẩn Phần 2 (thứ tự mặc định); người dùng kéo đổi thứ tự trong tab Cấu hình cột
// 9 màu cơ bản cho màu chữ / tô nền ô
const BASIC_COLORS: [string, string][] = [['#FF0000', 'Đỏ'], ['#00FF00', 'Xanh lá cây'], ['#0000FF', 'Xanh dương'], ['#FFFF00', 'Vàng'], ['#00FFFF', 'Lục lam (Cyan)'], ['#FF00FF', 'Hồng (Magenta)'], ['#FFFFFF', 'Trắng'], ['#000000', 'Đen'], ['#808080', 'Xám']];
const P2_STD: { key: string; label: string; desc: string; edit?: boolean; title?: string }[] = [
  { key: 'qty', label: 'KL', desc: 'Khối lượng soát xét – mặc định theo NTP, sửa được', edit: true },
  { key: 'vt', label: 'ĐG VT', desc: 'Đơn giá vật tư soát xét, sửa được', edit: true },
  { key: 'nc', label: 'ĐG NC', desc: 'Đơn giá nhân công soát xét, sửa được', edit: true },
  { key: 'dep', label: '% KH', desc: '% khấu hao vật tư, sửa được (20% hoặc 0,2; để trống = 100%)', edit: true },
  { key: 'total', label: 'ĐG tổng', desc: 'Tự tính = ĐG VT × %KH + ĐG NC' },
  { key: 'amt', label: 'Thành tiền', desc: 'Tự tính = ĐG tổng × KL' },
  { key: 'diff', label: 'Chênh lệch', desc: 'Tự tính = Thành tiền soát xét − Thành tiền NTP' },
  { key: 'pct', label: '% Chênh lệch', desc: 'Tự tính = Chênh lệch / Thành tiền NTP (▲ tăng, ▼ giảm)', title: '% chênh lệch = Chênh lệch / Thành tiền NTP' },
  { key: 'status', label: 'Kết quả', desc: 'NTP cao hơn: Cao (đỏ) · thấp hơn: Thấp (xanh dương) · bằng: Hợp lý (xanh lá)' },
  { key: 'note', label: 'Ghi chú', desc: 'Ghi chú soát xét', edit: true },
];
const P2_W: Record<string, number> = Object.fromEntries(P2_STD.map((c, i) => [c.key, P2_WIDTHS[i]]));
// PA3 – Phần 3: P.KSCP soát xét = các cột Phần 2 lặp lại lần 2 (ô chưa sửa = theo Phần 2 BCHCT)
const K3_STD = P2_STD.map(c => ({ ...c, key: 'k3_' + c.key, desc: c.desc.replace('soát xét', 'P.KSCP soát xét') + ' – trống = theo Phần 2 (BCHCT)' }));
const K3_KEYS = K3_STD.map(c => c.key);
K3_KEYS.forEach(k => { P2_W[k] = P2_W[k.slice(3)]; });
const stdOf = (k: string) => P2_STD.find(x => x.key === k) || K3_STD.find(x => x.key === k);
// Ô Phần 3 sửa được: khóa cột → trường dữ liệu (lưu rv_extra "k3_*")
const K3_FIELD: Record<string, string> = { k3_qty: 'x_k3_qty', k3_vt: 'x_k3_vt', k3_nc: 'x_k3_nc', k3_dep: 'x_k3_dep', k3_total: 'x_k3_total', k3_amt: 'x_k3_amt', k3_note: 'x_k3_note' };
const K3_NUMF = new Set(['x_k3_qty', 'x_k3_vt', 'x_k3_nc', 'x_k3_dep', 'x_k3_total', 'x_k3_amt']);
// Trường ô → khóa cột (địa chỉ ô, định dạng ô, công thức sẵn)
const FIELD_KEY: Record<string, string> = { rv_quantity: 'qty', rv_material: 'vt', rv_labor: 'nc', rv_dep_pct: 'dep', rv_total: 'total', x_ov_rv_amt: 'amt',
  ...Object.fromEntries(Object.entries(K3_FIELD).map(([k, f]) => [f, k])) };
const BUILTIN_OF: Record<string, string> = { rv_total: 'total', x_ov_rv_amt: 'amt', x_k3_total: 'k3_total', x_k3_amt: 'k3_amt', x_k3_qty: 'k3_qty', x_k3_vt: 'k3_vt', x_k3_nc: 'k3_nc', x_k3_dep: 'k3_dep' };
// Tiêu đề khối theo phương án
const BAND: Record<string, [string, string, string]> = {
  PA1: ['PHẦN 1: DỰ TOÁN ĐNTC', '', ''],
  PA2: ['PHẦN 1: BÁO GIÁ NTP', 'PHẦN 2: GIÁ BCHCT SOÁT XÉT', ''],
  PA3: ['PHẦN 1: BÁO GIÁ NTP', 'PHẦN 2: GIÁ BCHCT SOÁT XÉT', 'PHẦN 3: P.KSCP SOÁT XÉT'],
};
function p2Order(order: string[] | undefined, xcols: CustomCol[]) {
  const valid = [...P2_STD.map(c => c.key), ...xcols.map(c => 'x_' + c.id)];
  const out = (order || []).map(k => (k === 'lib' ? 'pct' : k)).filter((k, i, a) => valid.includes(k) && a.indexOf(k) === i);
  valid.forEach(k => { if (!out.includes(k)) out.push(k); });
  return out;
}
const FIELD_WIDTH: Record<string, number> = { stt: 56, code: 84, description: 260, specification: 140, brand: 100, model: 90, unit: 46,
  quantity: 70, material_price: 88, labor_price: 84, depreciation_pct: 64, dep_done_pct: 64, start_date: 92, end_date: 92, machine_price: 74, price: 92, amount: 108, note: 120 };
// Cột bắt buộc của file NTP (form chuẩn): STT, mô tả, ĐV, KL, ĐG VT, ĐG NC, %KH, thời gian bắt đầu / kết thúc, đơn giá, thành tiền, ghi chú
const REQUIRED_COLS: [string, string][] = [['stt', 'STT'], ['description', 'Mô tả công việc'], ['unit', 'Đơn vị'], ['quantity', 'Khối lượng'],
  ['material_price', 'Đơn giá vật tư'], ['labor_price', 'Đơn giá nhân công'], ['depreciation_pct', '% KH vật tư'], ['start_date', 'Thời gian bắt đầu'],
  ['end_date', 'Thời gian kết thúc'], ['price', 'Đơn giá'], ['amount', 'Thành tiền'], ['note', 'Ghi chú']];
// Ngày → dd/mm/yyyy (ô ngày Excel, chuỗi ISO, số ngày Excel); chuỗi đã là dd/mm/yyyy giữ nguyên
const dateText = (v: any): string => {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' && v > 20000 && v < 80000) { const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000); return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`; }
  const t = String(v).trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t); if (m) return `${m[3].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[1]}`;
  m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/.exec(t); if (m) return `${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[3].length === 2 ? '20' + m[3] : m[3]}`;
  return t;
};
const NUM_FIELDS = new Set(['quantity', 'material_price', 'labor_price', 'depreciation_pct', 'dep_done_pct', 'machine_price', 'price', 'amount']);
// % khấu hao vật tư: 0,2 → "20%"
const pctText = (v: number | null | undefined) => v === null || v === undefined ? '' : `${new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 }).format(v * 100)}%`;
// Số lẻ nhỏ (hệ số, tỷ lệ…) giữ phần thập phân; số tiền làm tròn đồng. Ô % đã được backend đổi sẵn thành chuỗi "5%".
const fmtCell = (v: any, field?: string) => (v === null || v === undefined ? '' : field === 'start_date' || field === 'end_date' ? dateText(v) : typeof v === 'number'
  ? field === 'depreciation_pct' || field === 'dep_done_pct' ? pctText(v > 1 ? v / 100 : v)
  : (field === 'quantity' || (!Number.isInteger(v) && Math.abs(v) < 1000) ? num(v) : money(v)) : String(v));
// Cột Phần 1 khi gói soát xét từ bản cũ (chưa chụp cấu trúc file).
const LEGACY_P1: P1Col[] = [
  { key: 'stt', label: 'STT', field: 'stt', width: 56, cls: 'link', get: (_r, it) => it?.boq_stt || '' },
  { key: 'desc', label: 'Hạng mục', field: 'description', width: 280, cls: 'desc', get: (_r, it) => it?.boq_description },
  { key: 'unit', label: 'ĐVT', field: 'unit', width: 52, cls: 'c', get: (_r, it) => it?.boq_unit },
  { key: 'qty', label: 'Khối lượng', field: 'quantity', width: 84, cls: 'r', get: (_r, it) => num(it?.boq_quantity) },
  { key: 'vt', label: 'Đơn giá VT', field: 'material_price', width: 92, cls: 'r', get: (_r, it) => money(it?.boq_material) },
  { key: 'nc', label: 'Đơn giá NC', field: 'labor_price', width: 92, cls: 'r', get: (_r, it) => money(it?.boq_labor) },
  { key: 'kh', label: 'Khấu hao', field: 'machine_price', width: 84, cls: 'r', get: (_r, it) => money(it?.boq_machine) },
  { key: 'tong', label: 'Đơn giá tổng', field: 'price', width: 98, cls: 'r', get: (_r, it) => money(it?.boq_price) },
  { key: 'tt', label: 'Thành tiền', field: 'amount', width: 116, cls: 'r', get: (_r, it) => money((it?.boq_quantity || 0) * (it?.boq_price || 0)) },
  { key: 'note', label: 'Ghi chú', field: 'note', width: 140, cls: 'ntp-note', get: (_r, it) => it?.boq_note || '' },
];
const PRICE_FIELDS = new Set(['rv_material', 'rv_labor', 'rv_total', 'rv_machine', 'x_k3_vt', 'x_k3_nc', 'x_k3_total']);
// Cột NTP (Phần 1) sửa được / gõ công thức được: trường cấu hình → trường dữ liệu
const NTP_EDIT: Record<string, 'boq_quantity' | 'boq_material' | 'boq_labor' | 'boq_dep_pct' | 'boq_price' | 'boq_dep_done'> = {
  quantity: 'boq_quantity', material_price: 'boq_material', labor_price: 'boq_labor', depreciation_pct: 'boq_dep_pct', price: 'boq_price', dep_done_pct: 'boq_dep_done' };
const NTP_VAL: Record<string, 'q' | 'm' | 'l' | 'd' | 'p' | 'k'> = { boq_quantity: 'q', boq_material: 'm', boq_labor: 'l', boq_dep_pct: 'd', boq_price: 'p', boq_dep_done: 'k' };
const PARTS: (keyof RV)[] = ['rv_material', 'rv_labor', 'rv_machine'];

// ru: làm tròn đơn giá (100 = …00). Mỗi đơn giá và đơn giá tổng đều làm tròn – giống backend / Excel ROUND(…,-2)
function calc(it: Item, e?: Edit, ru = 0) {
  const rv: RV = { rv_quantity: it.rv_quantity, rv_material: it.rv_material, rv_labor: it.rv_labor, rv_machine: it.rv_machine, rv_total: it.rv_total, rv_note: it.rv_note, rv_dep_pct: it.rv_dep_pct, ...e };
  const rd = (v: number | null | undefined) => (v === null || v === undefined ? (v ?? null) : roundUnit(v, ru));
  rv.rv_material = rd(rv.rv_material); rv.rv_labor = rd(rv.rv_labor); rv.rv_machine = rd(rv.rv_machine);
  const dep = rv.rv_dep_pct ?? ((e as any)?.boq_dep_pct !== undefined ? (e as any).boq_dep_pct : it.boq_dep_pct) ?? null;
  const hasParts = PARTS.some(k => rv[k] !== null && rv[k] !== undefined);
  // Số liệu NTP (Phần 1) – có thể đã sửa trên bảng
  const ev = (k: string) => ((e as any)?.[k] !== undefined ? (e as any)[k] : (it as any)[k]) as number | null;
  const ntp = { q: ev('boq_quantity'), m: ev('boq_material'), l: ev('boq_labor'), d: ev('boq_dep_pct'), p: ev('boq_price'), k: ev('boq_dep_done') };
  // Có % ĐÃ KH (vật tư dùng lại) → ĐG NC soát xét bị khóa, không tính nhân công
  const depDone = !!ntp.k;
  if (depDone) rv.rv_labor = null;
  if (e && ['boq_material', 'boq_labor', 'boq_dep_pct'].some(k => (e as any)[k] !== undefined) && (e as any).boq_price === undefined && (ntp.m !== null || ntp.l !== null))
    ntp.p = (ntp.m ?? 0) * (ntp.d ?? 1) + (ntp.l ?? 0) + (it.boq_machine ?? 0);
  const ntpTotal = ntp.p || 0;
  // Thành tiền gõ / công thức riêng (như ô Excel độc lập) – trống = KL × Đơn giá
  const ovN = ev('x_ov_ntp_amt'), ovR = ev('x_ov_rv_amt');
  const ntpAmt = typeof ovN === 'number' ? ovN : (ntp.q || 0) * ntpTotal;
  // Có làm tròn: KL soát xét làm tròn số nguyên (như Excel ROUND(…,0))
  const qty = ru ? roundUnit(rv.rv_quantity ?? ntp.q ?? 0, 1) : rv.rv_quantity ?? ntp.q ?? 0;
  // Có % KH vật tư: Đơn giá = VT × %KH + NC + Khấu hao
  // ĐG tổng gõ / công thức riêng (như ô Excel độc lập) được ưu tiên; không thì = VT × %KH + NC
  const total = roundUnit(rv.rv_total !== null && rv.rv_total !== undefined ? rv.rv_total : hasParts ? (rv.rv_material ?? 0) * (dep ?? 1) + (rv.rv_labor ?? 0) + (rv.rv_machine ?? 0) : ntpTotal, ru);
  const amt = typeof ovR === 'number' ? ovR : qty * total;
  const status = amt < ntpAmt - 0.5 ? 'Cao' : amt > ntpAmt + 0.5 ? 'Thấp' : 'Hợp lý';
  const st2 = it.is_added ? 'Bổ sung' : status;
  // Phần 3 (PA3 – P.KSCP): ô chưa sửa lấy theo Phần 2
  const xv = (k: string) => { const v = (e as any)?.['x_' + k] !== undefined ? (e as any)['x_' + k] : (it as any)['x_' + k]; return typeof v === 'number' ? v : null; };
  const q3 = xv('k3_qty'); const qty3 = q3 === null ? qty : ru ? roundUnit(q3, 1) : q3;
  const vt3 = xv('k3_vt') ?? rv.rv_material; const nc3 = depDone ? null : xv('k3_nc') ?? rv.rv_labor; const dep3 = xv('k3_dep') ?? dep;
  const parts3 = [vt3, nc3, rv.rv_machine].some(v => v !== null && v !== undefined);
  const touched3 = ['k3_vt', 'k3_nc', 'k3_dep'].some(k => xv(k) !== null);
  const t3 = xv('k3_total');
  const total3 = t3 !== null ? roundUnit(t3, ru) : parts3 && ((rv.rv_total === null || rv.rv_total === undefined) || touched3)
    ? roundUnit((rd(vt3) ?? 0) * (dep3 ?? 1) + (rd(nc3) ?? 0) + (rv.rv_machine ?? 0), ru) : total;
  const amt3 = xv('k3_amt') ?? qty3 * total3;
  const s3 = it.is_added ? 'Bổ sung' : amt3 < ntpAmt - 0.5 ? 'Cao' : amt3 > ntpAmt + 0.5 ? 'Thấp' : 'Hợp lý';
  const k3 = { qty: qty3, vt: vt3, nc: nc3, dep: dep3, total: total3, amt: amt3, diff: amt3 - ntpAmt, status: s3, parts: parts3 };
  return { rv, dep, ntp, depDone, hasParts, ntpTotal, ntpAmt, qty, total, amt, diff: amt - ntpAmt, status: st2, k3 };
}
// Công thức hiển thị trên thanh fx khi bấm vào ô (giống Excel)
type Calc = ReturnType<typeof calc>;
// Định dạng số theo ô (thanh công cụ)
const nfText = (v: number | null | undefined, nf?: CellStyle['nf']) => v === null || v === undefined || Number.isNaN(v) ? '' :
  nf === 'int' ? money(Math.round(v)) : nf === 'd2' ? new Intl.NumberFormat('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v)
    : nf === 'pct' ? `${new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 1 }).format(v * 100)}%` : null;
// rv_extra {id: giá trị, f_…: công thức, o_…: số NTP gốc} → it.x_<id>, it.f_…
// Như Excel: công thức bắt đầu bằng =, + hoặc - (khi có địa chỉ ô / hàm, vd -E7*2)
const isFormulaText = (s: string) => /^[=+]/.test(s) || (/^-/.test(s) && /[A-Za-z(\[]/.test(s));
const toFormula = (s: string) => (s.startsWith('=') ? s : '=' + (s.startsWith('+') ? s.slice(1) : s));
function spreadExtra(it: Item) {
  try { Object.entries(JSON.parse(it.rv_extra || '{}')).forEach(([k, v]) => { if (k === 'st') it.st = v as any; else if (!k.startsWith('o_')) (it as any)[k.startsWith('f_') ? k : 'x_' + k] = v; }); } catch { /* bỏ qua */ }
  if (it.is_added) { it.add_stt = it.boq_stt; it.add_desc = it.boq_description; it.add_unit = it.boq_unit; }
}
type SaveVer = { id: number; no: number; saved_at: string; saved_by: string; changes: number; note: string; ntp_total: number; review_total: number; has_file?: number; data?: Record<string, any> };
function formulaOf(it: Item, c: Calc, key: string): { name: string; expr: string } | null {
  const m = (v: number | null | undefined) => money(v ?? 0) || '0';
  const pt = (v: number | null | undefined) => (v === null || v === undefined ? '' : `${new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 }).format(v * 100)}%`);
  const unitExpr = (vt: number | null | undefined, dep: number | null | undefined, nc: number | null | undefined, kh: number | null | undefined) => {
    const names = [dep !== null && dep !== undefined ? 'ĐG VT × %KH' : 'ĐG VT', 'ĐG NC'];
    const vals = [dep !== null && dep !== undefined ? `${m(vt)} × ${pt(dep)}` : m(vt), m(nc)];
    if (kh) { names.push('Khấu hao'); vals.push(m(kh)); }
    return `${names.join(' + ')} = ${vals.join(' + ')}`;
  };
  const ntpParts = [it.boq_material, it.boq_labor, it.boq_machine].some(v => v !== null && v !== undefined);
  switch (key) {
    case 'total':
      return { name: 'Đơn giá soát xét', expr: c.hasParts ? `= ${unitExpr(c.rv.rv_material, c.dep, c.rv.rv_labor, c.rv.rv_machine)} = ${m(c.total)}` : `= ${m(c.total)} (nhập trực tiếp)` };
    case 'amt':
      return { name: 'Thành tiền soát xét', expr: `= Đơn giá × Khối lượng = ${m(c.total)} × ${num(c.qty)} = ${m(c.amt)}` };
    case 'diff':
      return { name: 'Chênh lệch', expr: `= Thành tiền soát xét − Thành tiền NTP = ${m(c.amt)} − ${m(c.ntpAmt)} = ${m(c.diff)}` };
    case 'price':
      return { name: 'Đơn giá NTP', expr: ntpParts ? `= ${unitExpr(it.boq_material, it.boq_dep_pct, it.boq_labor, it.boq_machine)} = ${m(c.ntpTotal)}` : `= ${m(c.ntpTotal)}` };
    case 'amount':
      return { name: 'Thành tiền NTP', expr: `= Đơn giá × Khối lượng = ${m(c.ntpTotal)} × ${num(it.boq_quantity || 0)} = ${m(c.ntpAmt)}` };
    case 'rv_quantity': return { name: 'Khối lượng soát xét', expr: num(c.qty) };
    case 'rv_material': return { name: 'Đơn giá VT soát xét', expr: m(c.rv.rv_material) };
    case 'rv_dep_pct': return { name: '% khấu hao vật tư', expr: pt(c.dep) };
    case 'rv_labor': return { name: 'Đơn giá NC soát xét', expr: m(c.rv.rv_labor) };
    case 'rv_machine': return { name: 'Khấu hao soát xét', expr: m(c.rv.rv_machine) };
    default: return null;
  }
}
// Nguồn giá cho nút Chạy
const SOURCES: [string, string, string][] = [
  ['default', 'Mặc định', 'ĐNTC chuẩn làm cơ sở; Báo giá NCC nạp sau ĐNTC thì ghi đè; vẫn chưa có giá thì lấy Hợp đồng tham khảo'],
  ['dntc', 'Đơn giá ĐNTC chuẩn', 'Lấy giá từ bảng đơn giá ĐNTC chuẩn'],
  ['hd', 'Hợp đồng tham khảo', 'Lấy giá từ hợp đồng tham khảo'],
  ['ncc', 'Báo giá NCC', 'So đầu mục và áp giá theo báo giá NCC'],
  ['kh', 'Khấu hao', '%KH theo thời gian: ROUNDUP((Kết thúc − Bắt đầu)/30) → bảng Khấu hao (Xa / Gần biển)'],
  ['truc', 'Đơn giá nhân công trực', 'Lấy giá từ bảng đơn giá nhân công trực'],
  ['ntp', 'Giữ giá NTP', 'Trả cột soát xét về đúng số liệu NTP'],
];
const libParts = (it: Item) => [it.lib_material, it.lib_labor, it.lib_machine].some(v => v !== null && v !== undefined);

export function PackageDetail({ id, onExit }: { id: number; onExit: () => void }) {
  const [d, setD] = useState<Detail | null>(null);
  const [versions, setVersions] = useState<Version[]>([]);
  const [edits, setEdits] = useState<Record<number, Edit>>({});
  const [form, setForm] = useState({ name: '', subcontractor: '', deadline: '', note: '', vat_percent: 10, status: '', version_id: 0, he_so: '1' });
  const [formDirty, setFormDirty] = useState(false);
  const [tab, setTab] = useState<'detail' | 'mapping' | 'history'>('detail');
  const [sel, setSel] = useState<Set<number>>(new Set());
  const [cur, setCur] = useState<number | null>(null);
  const [curSid, setCurSid] = useState<number | null>(null);
  const [menu, setMenu] = useState<{ x: number; y: number; sid: number; it?: Item } | null>(null);
  // Menu chuột phải trên chữ cột: chèn cột soát xét bên trái / phải, xóa cột tự tạo
  const [colMenu, setColMenu] = useState<{ x: number; y: number; at: number; key?: string; part1?: boolean } | null>(null);
  const [newCol, setNewCol] = useState<{ label: string; type: 'number' | 'text' }>({ label: '', type: 'number' });
  const [fx, setFx] = useState<{ ref: string; name: string; expr: string } | null>(null);
  // Ô đang chọn để sửa trên thanh công thức fx
  const [active, setActive] = useState<{ it: Item; field: string; pct?: boolean; addr: string } | null>(null);
  const [fxText, setFxText] = useState('');
  // Loại nội dung ô đang chọn: Công thức / Giá trị / Tự tính
  const [cellKind, setCellKind] = useState<'f' | 'v' | 'a' | null>(null);
  const [q, setQ] = useState('');
  const [filter, setFilter] = useState('');
  const [sheet, setSheet] = useState('');
  const [busy, setBusy] = useState('');
  const fileRef = useRef<HTMLInputElement>(null);
  const draftKey = d?.session_id ? `cx-draft-${id}-${d.session_id}` : '';

  const load = async (restoreDraft = false) => {
    const [p, v] = await Promise.all([mepApi<Detail>(`/api/mep/packages/${id}`), mepApi<Version[]>('/api/mep/library/versions?include_reference=true')]);
    (p.items as Item[]).forEach(spreadExtra);
    setD(p); setVersions(v); setSel(new Set()); setFormDirty(false);
    setForm({ name: p.name, subcontractor: p.subcontractor || '', deadline: p.deadline || '', note: p.note || '', vat_percent: p.vat_percent ?? 10, status: p.status, version_id: p.version_id || 0, he_so: String((p as any).he_so ?? 1).replace('.', ',') });
    let restored: Record<number, Edit> = {};
    if (restoreDraft && p.session_id) {
      try {
        const raw = localStorage.getItem(`cx-draft-${id}-${p.session_id}`);
        if (raw) { restored = JSON.parse(raw); const n = Object.keys(restored).length; if (n) toast(`Khôi phục ${n} dòng sửa chưa lưu từ lần trước – bấm Lưu để ghi`); }
      } catch { /* bỏ qua bản nháp hỏng */ }
    }
    setEdits(restored);
    if (!p.session_id && p.workbook_id) setTab('mapping');
  };
  useEffect(() => { load(true).catch(e => toast(e.message, 'err')); }, [id]);
  // Tự lưu nháp các ô đang sửa vào trình duyệt để không mất khi lỡ đóng trang.
  useEffect(() => {
    if (!draftKey) return;
    try { Object.keys(edits).length ? localStorage.setItem(draftKey, JSON.stringify(edits)) : localStorage.removeItem(draftKey); } catch { /* bỏ qua */ }
  }, [edits, draftKey]);

  // Thao tác chờ ghi lịch sử khi bấm Lưu (thêm / xóa dòng)
  const [pendingEv, setPendingEv] = useState<string[]>([]);
  const dirty = formDirty || Object.keys(edits).length > 0 || pendingEv.length > 0;
  useEffect(() => {
    const h = (e: BeforeUnloadEvent) => { if (dirty) { e.preventDefault(); e.returnValue = ''; } };
    window.addEventListener('beforeunload', h);
    return () => window.removeEventListener('beforeunload', h);
  }, [dirty]);

  // Xem lại một lần lưu cũ (chỉ xem): ghép số liệu bản lưu lên các dòng hiện tại
  const [viewVer, setViewVer] = useState<SaveVer | null>(null);
  const items = useMemo(() => {
    const base = d?.items || [];
    if (!viewVer?.data) return base;
    return base.map(it => {
      const snap = viewVer.data![String(it.id)];
      if (!snap) return it;
      const n = { ...it, ...snap } as Item;
      Object.keys(n).forEach(k => { if (k.startsWith('x_') || k.startsWith('f_') || k === 'st') delete (n as any)[k]; });
      spreadExtra(n);
      return n;
    });
  }, [d, viewVer]);
  const ru = d?.round_unit ?? 100;
  const pa: string = (d as any)?.pa || 'PA2';
  // Tính lại chỉ dòng thay đổi; dòng không đổi giữ nguyên đối tượng (để bảng không vẽ lại dòng đó)
  const calcCache = useRef(new Map<number, { it: Item; e: Edit | undefined; ru: number; c: ReturnType<typeof calc> }>());
  const calcs = useMemo(() => {
    const m = new Map<number, ReturnType<typeof calc>>(); const next = new Map<number, { it: Item; e: Edit | undefined; ru: number; c: ReturnType<typeof calc> }>();
    for (const it of items) {
      const prev = calcCache.current.get(it.id);
      const c = prev && prev.it === it && prev.e === edits[it.id] && prev.ru === ru ? prev.c : calc(it, edits[it.id], ru);
      m.set(it.id, c); next.set(it.id, { it, e: edits[it.id], ru, c });
    }
    calcCache.current = next;
    return m;
  }, [items, edits, ru]);
  const xcols = d?.custom_columns || [];
  // Giá trị cột tự tạo của từng dòng (số nhập tay hoặc kết quả công thức). Lỗi công thức → NaN.
  const refMap = useMemo(() => {
    const m = new Map<string, string>(REF_STD.map(([n, k]) => [refKey(n), k]));
    xcols.forEach(x => m.set(refKey(x.label), 'x_' + x.id));
    return m;
  }, [JSON.stringify(xcols)]);
  const rowEnv = (it: Item, c: ReturnType<typeof calc>, getX: (k: string) => number | null) => (name: string): number | null => {
    const k = refMap.get(refKey(name));
    if (!k) throw new Error(`Không có cột [${name}]`);
    switch (k) {
      case 'qty': return c.qty; case 'vt': return c.rv.rv_material; case 'nc': return c.rv.rv_labor; case 'dep': return c.dep ?? 1;
      case 'total': return c.total; case 'amt': return c.amt; case 'diff': return c.diff; case 'pct': return c.ntpAmt ? c.diff / c.ntpAmt : null; case 'lib': return it.lib_price;
      case 'n_qty': return it.boq_quantity; case 'n_vt': return it.boq_material; case 'n_nc': return it.boq_labor;
      case 'n_price': return c.ntpTotal; case 'n_amt': return c.ntpAmt;
      default: return getX(k);
    }
  };
  const xvals = useMemo(() => {
    const out = new Map<number, Record<string, number | null>>();
    const errs = new Map<string, string>();
    if (!xcols.some(x => x.type !== 'text')) return { vals: out, errs };
    for (const it of items) {
      const c = calcs.get(it.id)!;
      const memo: Record<string, number | null> = {};
      const busy = new Set<string>();
      const getX = (k: string): number | null => {
        if (k in memo) return memo[k];
        const x = xcols.find(y => 'x_' + y.id === k);
        if (!x || x.type === 'text') return null;
        if (x.type === 'number') {
          const e = (edits[it.id] as any)?.[k];
          const v = e !== undefined ? e : (it as any)[k];
          return (memo[k] = typeof v === 'number' ? v : v == null || v === '' ? null : parseVn(String(v)));
        }
        if (busy.has(k)) throw new Error('Công thức tham chiếu vòng');
        busy.add(k);
        try { memo[k] = evalFormula(x.formula || '', rowEnv(it, c, getX)); }
        catch (er: any) { memo[k] = NaN; if (!errs.has(k)) errs.set(k, er.message); }
        busy.delete(k);
        return memo[k];
      };
      xcols.forEach(x => { if (x.type !== 'text') getX('x_' + x.id); });
      out.set(it.id, memo);
    }
    return { vals: out, errs };
  }, [items, edits, calcs, refMap]);
  const xNumKeys = xcols.filter(x => x.type !== 'text').map(x => 'x_' + x.id);
  const sumX = (its: Item[]) => {
    const o: Record<string, number> = {};
    for (const k of xNumKeys) o[k] = its.reduce((s, it) => { const v = xvals.vals.get(it.id)?.[k]; return s + (typeof v === 'number' && !Number.isNaN(v) ? v : 0); }, 0);
    if (pa === 'PA3') { o.k3_amt = 0; o.k3_diff = 0; its.forEach(it => { const c = calcs.get(it.id); if (c) { o.k3_amt += c.k3.amt; o.k3_diff += c.k3.diff; } }); }
    return o;
  };
  const configs = d?.sheet_configs || [];
  const sheetNames = configs.length ? configs.map(c => c.sheet) : Array.from(new Set(items.map((i: any) => i.sheet_name as string)));
  const activeSheet = sheetNames.includes(sheet) ? sheet : sheetNames[0] || '';
  const cfg = configs.find(c => c.sheet === activeSheet);
  const structured = !!(cfg && cfg.first_col && d?.structure?.[activeSheet]);
  const itemById = useMemo(() => new Map(items.map(it => [it.id, it])), [items]);
  const matches = (it: Item) => {
    const c = calcs.get(it.id)!;
    if (filter === 'edited' && !(it.rv_touched || edits[it.id])) return false;
    if (filter.startsWith('lib:') && it.lib_flag !== filter.slice(4)) return false;
    if (filter && !filter.includes(':') && filter !== 'edited' && c.status !== filter) return false;
    if (q.trim()) return `${it.boq_stt} ${it.boq_description} ${it.lib_description || ''} ${c.rv.rv_note || ''} ${it.boq_note || ''}`.toLowerCase().includes(q.toLowerCase());
    return true;
  };
  const filtering = !!(filter || q.trim());
  // Dòng hiển thị theo đúng thứ tự file (gồm dòng nhóm/tổng); đang lọc thì chỉ còn dòng mục khớp.
  const viewRows: { r: SRow; it?: Item }[] = useMemo(() => {
    if (structured) {
      return d!.structure![activeSheet].map(r => ({ r, it: r.item_id ? itemById.get(r.item_id) : undefined }))
        .filter(x => (x.r.kind === 'item' ? !!x.it && matches(x.it) : !filtering));
    }
    return items.filter((it: any) => (it.sheet_name || activeSheet) === activeSheet && matches(it))
      .map(it => ({ r: { row: it.excel_row, kind: 'item' as const, cells: [], item_id: it.id }, it }));
  }, [d, items, calcs, q, filter, activeSheet, structured]);
  const rows = viewRows.filter(x => x.it).map(x => x.it!);
  const sheetItems = items.filter((it: any) => (it.sheet_name || activeSheet) === activeSheet);
  // Dòng nhóm / dòng tổng của file gốc có Thành tiền NTP → hiện tổng soát xét tương ứng ngay trên dòng đó.
  // Phạm vi: STT "3.3.5.4" gồm các dòng 3.3.5.4.x; dòng "TỔNG …" không STT gồm cả sheet.
  const groupTotals = useMemo(() => {
    const out = new Map<number, { rev: number; diff: number; ntp: number; all: boolean; xs: Record<string, number> }>();
    if (!structured || !cfg) return out;
    const idx = (f: string) => (cfg.mapping[f] ? cfg.mapping[f]! - cfg.first_col! : -1);
    const iStt = idx('stt'), iAmt = idx('amount'), iDesc = idx('description');
    const norm = (v: any) => String(v ?? '').trim().replace(/,/g, '.').replace(/\.+$/, '');
    const plain = (v: any) => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'd').toLowerCase();
    for (const r of d?.structure?.[activeSheet] || []) {
      if (r.kind !== 'group' || iAmt < 0 || typeof r.cells[iAmt] !== 'number') continue;
      const stt = iStt >= 0 ? norm(r.cells[iStt]) : '';
      const all = !stt && /\btong\b/.test(plain(iDesc >= 0 ? r.cells[iDesc] : '') + ' ' + plain(r.cells.join(' ')));
      if (!stt && !all) continue;
      let rev = 0, ntp = 0;
      const scope: Item[] = [];
      for (const it of sheetItems) {
        if (!all && !norm(it.boq_stt).startsWith(stt + '.')) continue;
        const c = calcs.get(it.id)!; rev += c.amt; ntp += c.ntpAmt; scope.push(it);
      }
      if (scope.length) out.set(r.row, { rev, ntp, diff: rev - ntp, all, xs: sumX(scope) });
    }
    return out;
  }, [d, cfg, structured, activeSheet, sheetItems, calcs, xvals]);
  const hasGrand = [...groupTotals.values()].some(g => g.all);
  const sheetTotals = useMemo(() => {
    let ntp = 0, rev = 0;
    for (const it of sheetItems) { const c = calcs.get(it.id)!; ntp += c.ntpAmt; rev += c.amt; }
    return { ntp, rev, diff: rev - ntp, pct: ntp ? ((rev - ntp) / ntp) * 100 : null, xs: sumX(sheetItems) };
  }, [sheetItems, calcs, xvals]);
  // Cột Phần 1 = cột gốc của sheet (đúng tiêu đề, thứ tự).
  // Độ rộng cột / chiều cao dòng người dùng kéo (nhớ theo gói + sheet trên máy này)
  const layoutKey = `cx-layout-${id}-${activeSheet}`;
  const [colW, setColW] = useState<Record<string, number>>({});
  const rowHRef = useRef<Record<string, number>>({});
  useEffect(() => {
    try { const v = JSON.parse(localStorage.getItem(layoutKey) || '{}'); setColW(v.c || {}); rowHRef.current = v.r || {}; } catch { setColW({}); rowHRef.current = {}; }
  }, [layoutKey]);
  const saveLayout = (c: Record<string, number>) => { try { localStorage.setItem(layoutKey, JSON.stringify({ c, r: rowHRef.current })); } catch { /* bỏ qua */ } };
  const p1: P1Col[] = useMemo(() => {
    if (!structured || !cfg) return LEGACY_P1.map(c => ({ ...c, width: colW['p1:' + c.key] ?? c.width }));
    const fieldOf: Record<number, string> = {};
    Object.entries(cfg.mapping).forEach(([f, c]) => { if (c) fieldOf[c] = f; });
    const out: P1Col[] = [];
    for (let col = cfg.first_col!; col <= cfg.last_col!; col++) {
      const f = fieldOf[col];
      const i = col - cfg.first_col!;
      out.push({ key: `c${col}`, label: cfg.headers[String(col)] || '', field: f, width: colW[`p1:c${col}`] ?? (f ? FIELD_WIDTH[f] || 110 : 110),
        cls: f === 'description' ? 'desc' : f === 'stt' ? 'link' : f === 'unit' || f === 'start_date' || f === 'end_date' ? 'c' : f === 'note' ? 'ntp-note' : f && NUM_FIELDS.has(f) ? 'r' : 'raw',
        get: r => fmtCell(r.cells[i], f) });
    }
    return out;
  }, [structured, cfg, colW]);
  const p2base = p2Order(d?.p2_order, xcols);
  // PA1 Dự toán ĐNTC: chỉ Phần 1 · PA3: thêm Phần 3 (P.KSCP) sau Phần 2
  const p2keys = pa === 'PA1' ? [] : pa === 'PA3' ? [...p2base, ...K3_KEYS] : p2base;
  const n3 = pa === 'PA3' ? K3_KEYS.length : 0;
  // Cột Phần 2 trong Excel xuất ra bắt đầu sau vùng bảng gốc 1 cột trống
  const p2Start = cfg?.p2_start || (structured && cfg?.last_col ? cfg.last_col : 10) + 2;
  const xById = new Map(xcols.map(x => ['x_' + x.id, x]));
  const p2w = p2keys.map(k => colW['p2:' + k] ?? P2_W[k] ?? (xById.get(k)?.type === 'number' ? 110 : 160));
  // Dòng tổng nhóm / tổng sheet ở Phần 2: chỉ Thành tiền, Chênh lệch, % (đặt ở cột Giá TV)
  // % chênh lệch = Chênh lệch / Thành tiền NTP, kèm tam giác ▲ tăng / ▼ giảm
  const pctCell = (k: string, v: number | null, strong?: boolean, title?: string) => (
    <td key={k} className={`r pctd ${strong ? 'strong' : ''} ${v !== null && v < -0.05 ? 'neg' : v !== null && v > 0.05 ? 'pos' : ''}`} title={title}>
      {v === null ? '' : Math.abs(v) < 0.05 ? '0%' : <>{v > 0 ? '▲' : '▼'} {new Intl.NumberFormat('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(Math.abs(v))}%</>}
    </td>);
  // Cột tự tạo dạng số / công thức: SUBTOTAL (cộng các dòng trong nhóm)
  const sumCells = (rev: number, diff: number, ntp: number, xs: Record<string, number> = {}) => p2keys.map(k =>
    k === 'k3_pct' ? pctCell(k, ntp && xs.k3_diff !== undefined ? (xs.k3_diff / ntp) * 100 : null, true)
    : k === 'k3_diff' ? <td key={k} className={`r ${(xs[k] || 0) < 0 ? 'neg' : (xs[k] || 0) > 0 ? 'pos' : ''}`}>{money(xs[k] || 0) || '0'}</td>
    : k in xs ? <td key={k} className="r strong" title="Tổng nhóm (SUBTOTAL)">{money(xs[k]) || '0'}</td>
    : k === 'amt' ? <td key={k} className="r strong">{money(rev)}</td>
      : k === 'diff' ? <td key={k} className={`r ${diff < 0 ? 'neg' : diff > 0 ? 'pos' : ''}`}>{money(diff) || '0'}</td>
        : k === 'pct' ? pctCell(k, ntp ? (diff / ntp) * 100 : null, true) : <td key={k} />);
  // Cố định cột #, ☐ và các cột gốc tới cột Hạng mục (nếu không quá rộng).
  const frozen = useMemo(() => {
    // Cố định tới hết cột Mô tả công việc + Đơn vị (nếu Đơn vị nằm ngay sau)
    const dsc = p1.findIndex(c => c.field === 'description');
    const di = dsc >= 0 && p1[dsc + 1]?.field === 'unit' ? dsc + 1 : dsc;
    let n = 0, w = 72;
    if (di >= 0) {
      const need = p1.slice(0, di + 1).reduce((a, c) => a + c.width, 0);
      if (need <= 640) { n = di + 1; w += need; }
    }
    const lefts = [0, 40]; let x = 72;
    p1.slice(0, n).forEach(c => { lefts.push(x); x += c.width; });
    return { n, width: w, lefts };
  }, [p1]);
  const stick = (i: number, head = false): React.CSSProperties | undefined =>
    i < 2 + frozen.n ? { position: 'sticky', left: frozen.lefts[i], zIndex: head ? 5 : 2, ...(i === 1 + frozen.n ? { boxShadow: '2px 0 0 #c4d0df' } : {}) } : undefined;
  const totals = useMemo(() => {
    let ntp = 0, rev = 0; const counts: Record<string, number> = { 'Hợp lý': 0, 'Cao': 0, 'Thấp': 0 };
    for (const it of items) { const c = calcs.get(it.id)!; const a = pa === 'PA3' ? c.k3.amt : c.amt; const s0 = pa === 'PA3' ? c.k3.status : c.status;
      ntp += c.ntpAmt; rev += a; counts[s0] = (counts[s0] || 0) + 1; }
    return { ntp, rev, diff: rev - ntp, pct: ntp ? ((rev - ntp) / ntp) * 100 : null, counts };
  }, [items, calcs, pa]);
  const libCounts = useMemo(() => {
    const c: Record<string, number> = {};
    items.forEach(it => { c[it.lib_flag] = (c[it.lib_flag] || 0) + 1; });
    c.edited = items.filter(it => it.rv_touched || edits[it.id]).length;
    return c;
  }, [items, edits]);

  const setF = (k: keyof typeof form, v: any) => { setForm(f => ({ ...f, [k]: v })); setFormDirty(true); };
  // Hoàn tác / làm lại (Ctrl+Z / Ctrl+Y) theo trạng thái ô chưa lưu
  const undoRef = useRef<{ past: Record<number, Edit>[]; future: Record<number, Edit>[]; prev: Record<number, Edit>; skip: boolean }>({ past: [], future: [], prev: {}, skip: false });
  useEffect(() => {
    const u = undoRef.current;
    if (u.skip) { u.skip = false; u.prev = edits; return; }
    if (u.prev !== edits) { u.past.push(u.prev); if (u.past.length > 100) u.past.shift(); u.future = []; u.prev = edits; }
  }, [edits]);
  const undoKeysRef = useRef<{ undo: () => void; redo: () => void }>({ undo: () => {}, redo: () => {} });
  const gridKeysRef = useRef<(e: KeyboardEvent) => boolean>(() => false);
  useEffect(() => {  // Ctrl+Z / Ctrl+Y khi đang ở thanh công cụ hoặc ngoài ô nhập
    const h = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest?.('.cx-grid.rv') || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || (t.tagName === 'INPUT' && !(t as HTMLInputElement).type.match(/color|checkbox/))) return;
      // Vùng chọn bằng số Dòng / chữ cột (không có ô đang gõ): Ctrl+C, Ctrl+V, Delete, Insert vẫn dùng được
      if (gridKeysRef.current(e)) return;
      if (!(e.ctrlKey || e.metaKey)) return;
      const k = e.key.toLowerCase(); if (k !== 'z' && k !== 'y') return;
      e.preventDefault(); k === 'z' ? undoKeysRef.current.undo() : undoKeysRef.current.redo();
    };
    window.addEventListener('keydown', h); return () => window.removeEventListener('keydown', h);
  }, []);
  const undo = () => { const u = undoRef.current; const p = u.past.pop(); if (!p) { toast('Không còn thao tác để hoàn tác'); return; } u.future.push(edits); u.skip = true; setEdits(p); };
  const redo = () => { const u = undoRef.current; const n = u.future.pop(); if (!n) return; u.past.push(edits); u.skip = true; setEdits(n); };
  undoKeysRef.current = { undo, redo };
  const editRow = (it: Item, patch: Edit) => viewVer ? toast('Đang xem bản lưu cũ – bấm "Về bản hiện tại" để sửa', 'err') : setEdits(e => {
    const next: Edit = { ...e[it.id], ...patch };
    // Bỏ các trường trùng giá trị đã lưu để "chưa lưu" phản ánh đúng.
    Object.keys(next).forEach(k => { if (((next as any)[k] ?? null) === ((it as any)[k] ?? null)) delete (next as any)[k]; });
    const out = { ...e };
    if (Object.keys(next).length) out[it.id] = next; else delete out[it.id];
    return out;
  });

  const save = async (silent = false) => {
    if (!d) return false;
    setBusy('save');
    try {
      if (formDirty) {
        await mepApi(`/api/mep/packages/${d.id}`, { method: 'PUT', body: JSON.stringify({
          name: form.name, subcontractor: form.subcontractor, deadline: form.deadline, note: form.note,
          vat_percent: form.vat_percent, status: form.status, version_id: form.version_id || null, he_so: form.he_so.replace(',', '.') || '1',
        }) });
      }
      const rowsToSave = Object.entries(edits).map(([rid, e]) => ({ id: Number(rid), ...e }));
      let saved = 0;
      if (rowsToSave.length || pendingEv.length) saved = (await mepApi<{ saved: number }>(`/api/mep/packages/${d.id}/save`, { method: 'POST', body: JSON.stringify({ rows: rowsToSave, events: pendingEv }) })).saved;
      setPendingEv([]);
      try { if (draftKey) localStorage.removeItem(draftKey); } catch { /* bỏ qua */ }
      if (!silent) toast(`Đã lưu${saved ? ` ${saved} dòng` : ''}`);
      await load();
      return true;
    } catch (e: any) { toast(e.message, 'err'); return false; } finally { setBusy(''); }
  };
  const exit = () => { if (!dirty || confirm('Có thay đổi chưa lưu. Thoát mà không lưu? (Các ô đã sửa vẫn được giữ nháp trên máy này.)')) onExit(); };

  const run = async (body: any = {}, confirmFirst = true) => {
    if (confirmFirst && Object.keys(edits).length && !confirm('Soát xét lại từ file sẽ bỏ các ô chưa lưu (các dòng đã lưu vẫn được giữ). Tiếp tục?')) return;
    setBusy('run');
    try {
      const r = await mepApi<any>(`/api/mep/packages/${id}/run`, { method: 'POST', body: JSON.stringify(body) });
      toast(`Đã soát xét ${r.summary.total} dòng · ${r.summary.higher} cao hơn thư viện · ${r.summary.not_found} không có trong thư viện${r.kept_edits ? ` · giữ ${r.kept_edits} dòng đã sửa` : ''}`);
      await load(); setTab('detail');
    } catch (e: any) { toast(`${e.message} – kiểm tra tab Cấu hình cột`, 'err'); await load(); setTab('mapping'); } finally { setBusy(''); }
  };
  const importFile = async (f: File) => {
    if (Object.keys(edits).length && !confirm('Nạp file mới sẽ bỏ các ô chưa lưu. Tiếp tục?')) return;
    setBusy('run');
    const fd = new FormData(); fd.append('file', f);
    try { await mepApi(`/api/mep/packages/${id}/file`, { method: 'POST', body: fd }); setEdits({}); setBusy(''); await run({}, false); }
    catch (e: any) { toast(e.message, 'err'); setBusy(''); }
  };
  const exportOriginal = async () => {
    if (dirty && !confirm('Có ô chưa lưu. Lưu trước rồi xuất?')) return;
    if (dirty && !(await save(true))) return;
    try {
      const warn = await mepDownload(`/api/mep/packages/${id}/export-original`);
      toast(warn === 'drawings' ? 'Đã xuất. Lưu ý: hình ảnh/biểu đồ trong file gốc có thể không được giữ – kiểm tra lại.' : 'Đã xuất file gốc kèm cột soát xét', warn ? 'err' : 'ok');
    } catch (e: any) { toast(e.message, 'err'); }
  };
  // ===== Thêm / chèn / xóa dòng KSCP bổ sung =====
  // Số liệu một dòng để chép sang dòng mới (hạng mục, ĐVT, KL, đơn giá, %KH, ghi chú, cột tự tạo)
  const rowPatchOf = (it: Item): Edit => {
    const c = calcs.get(it.id)!; const e = (edits[it.id] || {}) as any;
    const p: any = {
      add_stt: e.add_stt ?? it.boq_stt ?? null, add_desc: e.add_desc ?? it.boq_description ?? null, add_unit: e.add_unit ?? it.boq_unit ?? null,
      rv_quantity: c.qty, rv_material: c.rv.rv_material, rv_labor: c.rv.rv_labor, rv_dep_pct: c.dep, rv_note: c.rv.rv_note ?? null,
      rv_total: c.hasParts ? null : c.total,
    };
    xcols.forEach(x => { if (x.type !== 'formula') { const k = 'x_' + x.id; p[k] = e[k] !== undefined ? e[k] : (it as any)[k] ?? null; } });
    return p;
  };
  // Chèn các dòng đã copy xuống dưới dòng đang chọn (mặc định như Excel "Chèn ô đã sao chép")
  const insertCopiedRows = async (anchor?: number) => {
    const data = clipRef.current?.rowsData; if (!d || !data?.length) return;
    let sid = anchor ?? curSid; if (!sid) { toast('Bấm chọn một dòng trước', 'err'); return; }
    const created: { id: number; patch: Edit }[] = [];
    try {
      for (const patch of data) {
        const r: { id: number; sid: number } = await mepApi<{ id: number; sid: number }>(`/api/mep/packages/${d.id}/rows`, { method: 'POST', body: JSON.stringify({ anchor_sid: sid, where: 'after' }) });
        created.push({ id: r.id, patch }); sid = r.sid;
      }
    } catch (e: any) { toast(e.message, 'err'); }
    if (!created.length) return;
    setPendingEv(ev => [...ev, `Chèn ${created.length} dòng đã copy`]);
    await load(true);
    setEdits(prev => { const n = { ...prev }; created.forEach(x => { n[x.id] = { ...(n[x.id] || {}), ...x.patch }; }); return n; });
    setCur(created[0].id);
    toast(`Đã chèn ${created.length} dòng đã copy xuống dưới – bấm Lưu để ghi`);
  };
  const addRow = async (where: 'after' | 'before', anchor?: number) => {
    if (!d) return;
    const list = d.structure?.[activeSheet] || [];
    const sid = anchor ?? curSid ?? (where === 'after' ? list[list.length - 1]?.sid : list[0]?.sid);
    if (!sid) { toast('Bấm chọn một dòng trước', 'err'); return; }
    try {
      const r = await mepApi<{ id: number }>(`/api/mep/packages/${d.id}/rows`, { method: 'POST', body: JSON.stringify({ anchor_sid: sid, where }) });
      setPendingEv(ev => [...ev, where === 'after' ? 'Thêm 1 dòng KSCP bổ sung' : 'Chèn 1 dòng KSCP bổ sung']);
      await load(true);
      setCur(r.id);
      setTimeout(() => { const el = document.getElementById(`add-${r.id}`) as HTMLInputElement | null; el?.scrollIntoView({ block: 'center' }); el?.focus(); }, 80);
      toast(where === 'after' ? 'Đã thêm dòng – nhập hạng mục, KL, đơn giá rồi bấm Lưu' : 'Đã chèn dòng – nhập hạng mục, KL, đơn giá rồi bấm Lưu');
    } catch (e: any) { toast(e.message, 'err'); }
  };
  const delRow = async (it: Item) => {
    if (!d || !it.is_added) return;
    if (!confirm(`Xóa dòng "${it.boq_description || 'mới thêm'}"?`)) return;
    try {
      await mepApi(`/api/mep/packages/${d.id}/rows/${it.id}`, { method: 'DELETE' });
      setEdits(e => { const n = { ...e }; delete n[it.id]; return n; });
      setPendingEv(ev => [...ev, `Xóa dòng KSCP bổ sung: ${it.boq_description || '(chưa đặt tên)'}`]);
      setCur(null); await load(true); toast('Đã xóa dòng');
    } catch (e: any) { toast(e.message, 'err'); }
  };
  const saveColLayout = async (columns: CustomCol[], order: string[], msg: string) => {
    if (!d) return;
    try {
      await mepApi(`/api/mep/packages/${d.id}/columns`, { method: 'PUT', body: JSON.stringify({ columns, order, round_unit: d.round_unit ?? 100 }) });
      await load(true); toast(msg);
    } catch (e: any) { toast(e.message, 'err'); }
  };
  const insertCol = (side: 'left' | 'right') => {
    if (!colMenu) return;
    const label = newCol.label.trim();
    if (!label) { toast('Nhập tên cột', 'err'); return; }
    const pos = colMenu.part1 ? 0 : side === 'left' ? colMenu.at : colMenu.at + 1;
    const order = [...p2base]; order.splice(Math.min(pos, p2base.length), 0, `new:${xcols.length}`);
    setColMenu(null); setNewCol({ label: '', type: 'number' });
    saveColLayout([...xcols, { id: '', label, type: newCol.type }], order, `Đã thêm cột "${label}"`);
  };
  const deleteCol = (key: string) => {
    const x = xById.get(key); if (!x) return;
    if (!confirm(`Xóa cột "${x.label}"? Số liệu đã nhập ở cột này sẽ không hiện nữa.`)) return;
    setColMenu(null);
    saveColLayout(xcols.filter(c => 'x_' + c.id !== key), p2base.filter(k => k !== key), `Đã xóa cột "${x.label}"`);
  };
  useEffect(() => { if (!colMenu) return; const c = (e: MouseEvent) => { if (!(e.target as HTMLElement).closest?.('.bv-colmenu')) setColMenu(null); }; window.addEventListener('mousedown', c); return () => window.removeEventListener('mousedown', c); }, [colMenu]);
  useEffect(() => { if (!menu) return; const c = () => setMenu(null); window.addEventListener('click', c); window.addEventListener('scroll', c, true); return () => { window.removeEventListener('click', c); window.removeEventListener('scroll', c, true); }; }, [menu]);
  const targets = () => (sel.size ? items.filter(i => sel.has(i.id)) : cur ? items.filter(i => i.id === cur) : []);
  // Áp giá thư viện / Giữ giá NTP: nếu đang chọn VÙNG Ô ở phần soát xét thì chỉ áp đúng các cột / dòng được chọn
  type PriceX = { material: number | null; labor: number | null; machine: number | null; total: number; src?: string };
  const bulk = (mode: 'lib' | 'ntp' | 'kh', ext?: { prices?: Record<string, PriceX>; dep?: Record<string, number>; formula?: Record<string, string>; zero?: boolean; label?: string; all?: boolean }) => {
    // Từng vùng chọn (giữ Ctrl để chọn nhiều vùng) → các dòng + cột của vùng đó
    const jobs: { it: Item; keys: Set<string> | null; p3?: boolean }[] = [];
    const rs = allRanges();
    const PK = ['qty', 'vt', 'nc', 'dep', 'total'];
    const P1K: Record<string, string> = { material_price: 'vt', labor_price: 'nc', depreciation_pct: 'dep' };
    const keysOf = (g: { c1: number; c2: number }) => { const ks = new Set<string>(); for (let c = g.c1; c <= g.c2; c++) {
      const k0 = c - 2 < p1.length ? (pa === 'PA1' && mode !== 'ntp' ? P1K[p1[c - 2]?.field || ''] : undefined) : p2keys[c - 2 - p1.length];
      const k = k0?.startsWith('k3_') ? k0.slice(3) : k0;
      if (k && PK.includes(k)) ks.add(k); } return ks; };
    // PA3: vùng chọn ở Phần 3 (hoặc chưa chọn cột / cả bảng) → ghi vào Phần 3 P.KSCP; chọn ở Phần 2 → ghi Phần 2
    const in3 = (g: { c1: number; c2: number }) => pa === 'PA3' && (p2keys[g.c2 - 2 - p1.length] || '').startsWith('k3_');
    // Đang chọn ô / vùng ô trong phần soát xét → chỉ áp cho đúng các ô đó (ưu tiên hơn ô tích đầu dòng)
    const useRanges = rs.length && (!sel.size || rs.some(g => keysOf(g).size));
    if (useRanges) {
      for (const g of rs) {
        const ks = keysOf(g);
        rows.slice(g.r1, g.r2 + 1).forEach(it => jobs.push({ it, keys: ks.size ? ks : null, p3: ks.size ? in3(g) : pa === 'PA3' }));  // chọn cột khác (hạng mục…) → áp cả dòng
      }
    } else targets().forEach(it => jobs.push({ it, keys: null, p3: pa === 'PA3' }));
    if (!jobs.length && ext?.all) rows.forEach(it => jobs.push({ it, keys: null, p3: pa === 'PA3' }));  // chưa chọn gì → cả bảng
    if (!jobs.length) { toast('Chọn ô / vùng ô (hoặc ô vuông đầu dòng) trước', 'err'); return; }
    const allKeys = new Set<string>(); let anyRow = false;
    jobs.forEach(j => (j.keys ? j.keys.forEach(k => allKeys.add(k)) : (anyRow = true)));
    const patches = new Map<number, { it: Item; p: Edit }>();
    let n = 0;
    for (const { it, keys, p3 } of jobs) {
      if (it.is_added && mode === 'ntp') continue;
      const has = (k: string) => !keys || keys.has(k);
      const p: Edit = { ...(patches.get(it.id)?.p || {}) };
      // ĐG tổng: chọn kèm cả ĐG VT + ĐG NC (hoặc cả dòng) → để công thức tự tính; chọn riêng ô ĐG tổng → điền thẳng giá
      const totalByParts = !keys || (keys.has('vt') && keys.has('nc'));
      if (mode === 'kh') {  // %KH theo thời gian (bảng Khấu hao)
        const v = ext?.dep?.[String(it.id)];
        if (v !== undefined && has('dep')) { p.rv_dep_pct = v; const fm = ext?.formula?.[String(it.id)]; if (fm) (p as any).f_rv_dep_pct = fm; }
      } else if (mode === 'lib') {
        // Giá theo nguồn đã chọn (ext.prices); không có ext → giá thư viện đối chiếu sẵn của dòng
        const x: PriceX | undefined = ext?.prices ? ext.prices[String(it.id)]
          : it.lib_price ? { material: it.lib_material ?? null, labor: it.lib_labor ?? null, machine: it.lib_machine ?? null, total: it.lib_price } : undefined;
        if (!x && ext?.prices && !ext.zero) continue;  // nguồn không có giá dòng này → giữ nguyên
        const lp = x?.total || 0;  // không có giá → điền 0
        const parts = !!x && [x.material, x.labor, x.machine].some(v => v !== null && v !== undefined);
        if (has('vt')) p.rv_material = parts ? x!.material : 0;
        if (has('nc')) p.rv_labor = parts ? x!.labor : 0;
        if (!keys) p.rv_machine = parts ? x!.machine : null;
        if (has('total')) p.rv_total = totalByParts && parts ? null : lp;
        if (!keys && !parts) Object.assign(p, { rv_material: null, rv_labor: null, rv_total: lp });
      } else {
        const ntpParts = [it.boq_material, it.boq_labor, it.boq_machine].some(v => v !== null && v !== undefined);
        if (has('qty')) p.rv_quantity = it.boq_quantity ?? 0;
        if (has('dep')) p.rv_dep_pct = it.boq_dep_pct ?? null;
        if (has('vt')) p.rv_material = it.boq_material ?? null;
        if (has('nc')) p.rv_labor = it.boq_labor ?? null;
        if (!keys) p.rv_machine = it.boq_machine;
        if (has('total')) p.rv_total = totalByParts && ntpParts ? null : (it.boq_price ?? 0);
      }
      if (p3) {  // PA3: ghi vào cột Phần 3 (P.KSCP) thay vì Phần 2
        const M: Record<string, string> = { rv_quantity: 'x_k3_qty', rv_material: 'x_k3_vt', rv_labor: 'x_k3_nc', rv_dep_pct: 'x_k3_dep', rv_total: 'x_k3_total' };
        const q: any = {};
        Object.entries(p).forEach(([f, v]) => { if (M[f]) q[M[f]] = v; else if (f.startsWith('f_rv_') && M[f.slice(2)]) q['f_' + M[f.slice(2)]] = v; else if (f !== 'rv_machine') q[f] = v; });
        Object.keys(p).forEach(f => delete (p as any)[f]); Object.assign(p, q);
      }
      // Ô đang có công thức gõ tay → bỏ công thức để nhận giá vừa điền
      const FIELD: Record<string, string> = { rv_quantity: 'qty', rv_material: 'vt', rv_labor: 'nc', rv_dep_pct: 'dep', rv_total: 'total', x_k3_qty: 'qty', x_k3_vt: 'vt', x_k3_nc: 'nc', x_k3_dep: 'dep', x_k3_total: 'total' };
      Object.keys(p).forEach(f => { if (FIELD[f] && (p as any)['f_' + f] === undefined && formulaOfCell(it, f)) (p as any)['f_' + f] = null; });
      if (pa === 'PA1' && mode !== 'ntp') {  // Dự toán: giá / %KH ghi luôn vào cột E, F, H (Phần 1)
        if ('rv_material' in p) (p as any).boq_material = p.rv_material ?? 0;
        if ('rv_labor' in p) (p as any).boq_labor = p.rv_labor ?? 0;
        if ('rv_dep_pct' in p) { (p as any).boq_dep_pct = p.rv_dep_pct; if ((p as any).f_rv_dep_pct) (p as any).f_boq_dep_pct = (p as any).f_rv_dep_pct; }
        if ('rv_total' in p && p.rv_total !== null && !('rv_material' in p)) (p as any).boq_price = p.rv_total;
      }
      if (!Object.keys(p).length) continue;
      patches.set(it.id, { it, p });
    }
    patches.forEach(({ it, p }) => { editRow(it, p); n++; });
    const what = !anyRow && allKeys.size ? ` (chỉ cột ${[...allKeys].map(k => P2_STD.find(x => x.key === k)?.label).join(', ')})` : '';
    toast(mode === 'ntp' ? `Trả về giá NTP cho ${n} dòng${what} – bấm Lưu để ghi` : `${ext?.label || 'Áp giá thư viện'}: ${n} dòng${what} – bấm Lưu để ghi`);
  };
  // Nút Chạy: lấy giá / %KH theo nguồn đã chọn rồi điền vào ô đang chọn (chưa chọn gì → cả bảng)
  const [src, setSrc] = useState('default');
  // Báo giá NCC / Hợp đồng tham khảo: chọn file (bản thư viện) cụ thể – tên theo file đã nạp ở Thư viện giá
  const [libVers, setLibVers] = useState<{ id: number; name: string; kind?: string; note?: string; created_at?: string }[]>([]);
  const [srcVer, setSrcVer] = useState<number | ''>('');
  useEffect(() => {
    if (src !== 'ncc' && src !== 'hd') return;
    mepApi<any[]>('/api/mep/library/versions').then(vs => {
      const xs = vs.filter(v => (v.kind || 'dntc') === src); setLibVers(xs);
      setSrcVer(cur => (xs.some(x => x.id === cur) ? cur : xs[0]?.id ?? ''));
    }).catch(() => setLibVers([]));
  }, [src]);
  const verLabel = (v: { name: string; note?: string }) => (v.note || '').replace(/^Nạp từ\s*/, '') || v.name;
  const [running, setRunning] = useState(false);
  const runSource = async () => {
    if (src === 'ntp') { bulk('ntp', { all: true }); return; }
    const hasSel = allRanges().length > 0 || sel.size > 0;
    if (!hasSel && !confirm(`Chưa chọn ô nào – áp "${SOURCES.find(x => x[0] === src)?.[1]}" cho cả bảng?`)) return;
    setRunning(true);
    try {
      if ((src === 'ncc' || src === 'hd') && !srcVer) { toast(`Chưa có file ${src === 'ncc' ? 'Báo giá NCC' : 'Hợp đồng tham khảo'} – nạp ở Thư viện giá trước`, 'err'); return; }
      const r = await mepApi<any>(`/api/mep/packages/${id}/price-source`, { method: 'POST', body: JSON.stringify({ source: src, version_id: src === 'ncc' || src === 'hd' ? srcVer : undefined }) });
      const label = `${SOURCES.find(x => x[0] === src)?.[1]} (${(r.used || []).join(' · ')})`;
      if (src === 'kh') bulk('kh', { dep: r.dep || {}, formula: r.formula || {}, label, all: true });
      else bulk('lib', { prices: r.prices || {}, zero: src === 'default' || src === 'dntc', label, all: true });
    } catch (e: any) { toast(e.message, 'err'); } finally { setRunning(false); }
  };

  const hintKeyRef = useRef<(e: React.KeyboardEvent<HTMLInputElement>) => boolean>(() => false);
  const onCellKey = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (hintKeyRef.current(e)) return;
    const t = e.currentTarget; const r = Number(t.dataset.r); const c = t.dataset.c;
    const go = (dr: number) => {
      const el = document.querySelector<HTMLInputElement>(`.cx-grid input[data-c="${c}"][data-r="${r + dr}"]`);
      if (el) { e.preventDefault(); el.focus(); el.select(); }
    };
    if (e.key === 'Enter' || e.key === 'ArrowDown') go(1);
    else if (e.key === 'ArrowUp') go(-1);
    else if (e.key === 'Escape') t.blur();
  };
  // ===== Chọn vùng ô + copy / paste kiểu Excel =====
  // Vùng chọn tính theo (thứ tự dòng hạng mục đang hiện, chỉ số ô trong dòng). Kéo chuột hoặc Shift+bấm để chọn vùng.
  const gridRef = useRef<HTMLTableElement>(null);
  const [selA, setSelA] = useState<{ r: number; c: number } | null>(null);
  const [selB, setSelB] = useState<{ r: number; c: number } | null>(null);
  const dragging = useRef(false);
  const cellPos = (el: EventTarget | null) => {
    const td = (el as HTMLElement | null)?.closest?.('td');
    const tr = td?.parentElement as HTMLTableRowElement | undefined;
    if (!td || !tr || tr.dataset.ri === undefined) return null;
    return { r: Number(tr.dataset.ri), c: (td as HTMLTableCellElement).cellIndex, td: td as HTMLTableCellElement };
  };
  // Nhiều vùng chọn: giữ Ctrl + bấm / kéo để chọn thêm vùng (như Excel)
  type Rg = { r1: number; r2: number; c1: number; c2: number };
  const [extraR, setExtraR] = useState<Rg[]>([]);
  const allRanges = (): Rg[] => { const g = range(); return [...extraR, ...(g ? [g] : [])]; };
  const range = () => selA && selB ? { r1: Math.min(selA.r, selB.r), r2: Math.max(selA.r, selB.r), c1: Math.max(2, Math.min(selA.c, selB.c)), c2: Math.max(selA.c, selB.c) } : null;
  const rowEl = (ri: number) => gridRef.current?.querySelector<HTMLTableRowElement>(`tbody tr[data-ri="${ri}"]`);
  // Khóa thay đổi khi thân bảng vẽ lại (để các hiệu ứng DOM chạy lại đúng lúc)
  const bodyElKey = useMemo(() => ({}), [d, viewRows, calcs, edits, p1, p2keys.join('|'), cur, curSid, sel, xvals, activeSheet, viewVer]);
  useEffect(() => {  // tô vùng chọn (+ số dòng / chữ cột của vùng như Excel)
    const t = gridRef.current; if (!t) return;
    t.querySelectorAll('.csel, .csel-h, .csel1').forEach(e => { e.classList.remove('csel'); e.classList.remove('csel-h'); e.classList.remove('csel1'); });
    const rs = allRanges(); if (!rs.length) return;
    if (rs.length === 1 && rs[0].r1 === rs[0].r2 && rs[0].c1 === rs[0].c2) {  // 1 ô: khung xanh như Excel (cả ô chữ / ô không sửa được)
      const tr = rowEl(rs[0].r1); tr?.cells[rs[0].c1]?.classList.add('csel1');
      tr?.cells[0]?.classList.add('csel-h'); (t.querySelector('thead tr.letters') as HTMLTableRowElement | null)?.cells[rs[0].c1]?.classList.add('csel-h');
      return;
    }
    const lt = t.querySelector('thead tr.letters') as HTMLTableRowElement | null;
    for (const g of rs) {
      for (let r = g.r1; r <= g.r2; r++) { const tr = rowEl(r); if (!tr) continue; tr.cells[0]?.classList.add('csel-h'); for (let c = g.c1; c <= g.c2; c++) tr.cells[c]?.classList.add('csel'); }
      for (let c = g.c1; c <= g.c2; c++) lt?.cells[c]?.classList.add('csel-h');
    }
  }, [selA, selB, extraR, bodyElKey]);
  // Chọn cả dòng (bấm số Dòng) / cả cột (bấm chữ cột) – kéo để chọn nhiều
  const lineDrag = useRef<'row' | 'col' | null>(null);
  const lastCol = () => 1 + (gridRef.current?.querySelector('thead tr.letters') as HTMLTableRowElement | null)?.cells.length! - 2;
  // Kéo mép phải chữ cột để đổi độ rộng; kéo mép dưới số Dòng để đổi chiều cao (như Excel)
  const startColResize = (e: React.MouseEvent, ci: number) => {
    e.preventDefault(); e.stopPropagation();
    const t = gridRef.current; if (!t) return;
    const colEl = t.querySelectorAll('colgroup col')[ci] as HTMLElement | undefined; if (!colEl) return;
    const id = ci - 2 < p1.length ? `p1:${p1[ci - 2].key}` : `p2:${p2keys[ci - 2 - p1.length]}`;
    const x0 = e.clientX, w0 = colEl.getBoundingClientRect().width || parseFloat(colEl.style.width) || 80;
    const tw0 = t.getBoundingClientRect().width;
    document.body.classList.add('cx-resizing-col');
    const mv = (ev: MouseEvent) => { const w = Math.max(30, Math.round(w0 + ev.clientX - x0)); colEl.style.width = `${w}px`; t.style.minWidth = `${tw0 + (w - w0)}px`; };
    const up = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', mv); window.removeEventListener('mouseup', up); document.body.classList.remove('cx-resizing-col');
      const w = Math.max(30, Math.round(w0 + ev.clientX - x0));
      setColW(c => { const n = { ...c, [id]: w }; saveLayout(n); return n; });
    };
    window.addEventListener('mousemove', mv); window.addEventListener('mouseup', up);
  };
  const startRowResize = (e: React.MouseEvent) => {
    e.preventDefault(); e.stopPropagation();
    const tr = (e.target as HTMLElement).closest('tr') as HTMLTableRowElement | null; if (!tr) return;
    const key = tr.dataset.ri !== undefined ? `i${rows[Number(tr.dataset.ri)]?.id}` : `g${tr.dataset.sid}`;
    const y0 = e.clientY, h0 = tr.getBoundingClientRect().height;
    document.body.classList.add('cx-resizing-row');
    const mv = (ev: MouseEvent) => { tr.style.height = `${Math.max(16, Math.round(h0 + ev.clientY - y0))}px`; };
    const up = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', mv); window.removeEventListener('mouseup', up); document.body.classList.remove('cx-resizing-row');
      rowHRef.current = { ...rowHRef.current, [key]: Math.max(16, Math.round(h0 + ev.clientY - y0)) }; saveLayout(colW);
    };
    window.addEventListener('mousemove', mv); window.addEventListener('mouseup', up);
  };
  // Áp chiều cao dòng đã lưu sau mỗi lần vẽ bảng
  useEffect(() => {
    const t = gridRef.current; if (!t) return;
    Object.entries(rowHRef.current).forEach(([k, h]) => {
      const tr = k.startsWith('i') ? (() => { const ri = rows.findIndex(x => `i${x.id}` === k); return ri >= 0 ? rowEl(ri) : null; })() : t.querySelector<HTMLTableRowElement>(`tbody tr[data-sid="${k.slice(1)}"]`);
      if (tr) tr.style.height = `${h}px`;
    });
  });
  const onLetterDown = (ci: number, e: React.MouseEvent) => {
    if (e.button !== 0 || !rows.length) return;
    e.preventDefault(); (document.activeElement as HTMLElement | null)?.blur?.();
    const r2 = rows.length - 1;
    if (e.shiftKey && selA) { setSelA({ r: 0, c: selA.c }); setSelB({ r: r2, c: ci }); return; }
    setSelA({ r: 0, c: ci }); setSelB({ r: r2, c: ci }); lineDrag.current = 'col';
  };
  const onLetterEnter = (ci: number, e: React.MouseEvent) => { if (lineDrag.current === 'col' && (e.buttons & 1)) setSelB({ r: rows.length - 1, c: ci }); };
  // Ô vuông góc dưới-phải vùng chọn (fill handle) – kéo xuống để chép công thức / giá trị như Excel
  const fillRef = useRef<{ g: { r1: number; r2: number; c1: number; c2: number }; to: number } | null>(null);
  const fillApplyRef = useRef<() => void>(() => {});
  useEffect(() => {
    const t = gridRef.current; if (!t) return;
    t.querySelectorAll('.fill-h').forEach(e => e.remove());
    t.querySelectorAll('td.fh-host').forEach(e => e.classList.remove('fh-host'));
    const g = range(); if (!g) return;
    const td = rowEl(g.r2)?.cells[g.c2];
    if (!td || !td.querySelector('input:not([type=checkbox])')) return;
    const h = document.createElement('div'); h.className = 'fill-h'; h.title = 'Kéo xuống để chép công thức / giá trị (Ctrl+D)';
    if (!td.style.position) td.classList.add('fh-host');
    td.appendChild(h);
  }, [selA, selB, bodyElKey]);
  // Áp định dạng ô (đậm, nghiêng, gạch chân, màu chữ, màu nền, căn lề) lên các ô đang hiện
  useEffect(() => {
    const t = gridRef.current; if (!t) return;
    t.querySelectorAll('[data-sty]').forEach(el => {
      const h = el as HTMLElement; ['font-weight', 'font-style', 'text-decoration', 'color', 'background', 'text-align'].forEach(pr => h.style.removeProperty(pr));
      h.querySelectorAll('input').forEach(i => ['font-weight', 'font-style', 'text-decoration', 'color', 'text-align', 'background'].forEach(pr => i.style.removeProperty(pr)));
      h.removeAttribute('data-sty');
    });
    rows.forEach((it, ri) => {
      const st = stylesOf(it); if (!Object.keys(st).length) return;
      const tr = rowEl(ri); if (!tr) return;
      Object.entries(st).forEach(([id, sty]) => {
        const ci = id.startsWith('p1:') ? 2 + p1.findIndex(c => c.key === id.slice(3)) : 2 + p1.length + p2keys.indexOf(id.slice(3));
        const td = tr.cells[ci] as HTMLElement | undefined; if (!td || ci < 2) return;
        td.setAttribute('data-sty', '1');
        const set = (el: HTMLElement) => {
          if (sty.b) el.style.setProperty('font-weight', '700', 'important');
          if (sty.i) el.style.setProperty('font-style', 'italic', 'important');
          if (sty.u) el.style.setProperty('text-decoration', 'underline', 'important');
          if (sty.c) el.style.setProperty('color', sty.c, 'important');
          if (sty.al) el.style.setProperty('text-align', { l: 'left', c: 'center', r: 'right' }[sty.al], 'important');
        };
        set(td); td.querySelectorAll('input:not([type=checkbox])').forEach(i => set(i as HTMLElement));
        if (sty.bg) { td.style.setProperty('background', sty.bg, 'important'); td.querySelectorAll('input:not([type=checkbox])').forEach(i => (i as HTMLElement).style.setProperty('background', 'transparent', 'important')); }
      });
    });
  }, [bodyElKey]);
  fillApplyRef.current = () => {
    const f = fillRef.current; fillRef.current = null;
    gridRef.current?.querySelectorAll('.cfill').forEach(e => e.classList.remove('cfill'));
    if (!f || f.to <= f.g.r2) return;
    const h = f.g.r2 - f.g.r1 + 1; let n = 0;
    for (let c = f.g.c1; c <= f.g.c2; c++) for (let r = f.g.r2 + 1; r <= f.to; r++) n += putFromSource(r, c, srcCell(f.g.r1 + ((r - f.g.r1) % h), c));
    setSelA({ r: f.g.r1, c: f.g.c1 }); setSelB({ r: f.to, c: f.g.c2 });
    toast(n ? `Đã kéo công thức / giá trị xuống ${n} ô – bấm Lưu để ghi` : 'Không có ô sửa được để điền');
  };
  useEffect(() => { const up = () => { dragging.current = false; pointRef.current = null; lineDrag.current = null; if (fillRef.current) fillApplyRef.current(); }; window.addEventListener('mouseup', up); return () => window.removeEventListener('mouseup', up); }, []);
  // Chặn kéo-thả chữ giữa các ô (bấm đúp rồi kéo làm mất số ở ô cũ): chặn ở mức trình duyệt, pha capture
  useEffect(() => {
    const inGrid = (e: Event) => !!(e.target as HTMLElement | null)?.closest?.('.cx-grid.rv');
    const stop = (e: Event) => { if (inGrid(e)) { e.preventDefault(); e.stopPropagation(); } };
    const noDragInput = (e: Event) => { const t = (e as InputEvent).inputType || ''; if (inGrid(e) && /Drag|Drop/.test(t)) e.preventDefault(); };
    document.addEventListener('dragstart', stop, true);
    document.addEventListener('drop', stop, true);
    document.addEventListener('beforeinput', noDragInput, true);
    return () => { document.removeEventListener('dragstart', stop, true); document.removeEventListener('drop', stop, true); document.removeEventListener('beforeinput', noDragInput, true); };
  }, []);
  // ===== Chế độ trỏ ô như Excel: đang gõ công thức (=… ) mà bấm/kéo vào ô khác → chèn địa chỉ ô / vùng =====
  const pointRef = useRef<{ el: HTMLInputElement; start: number; end: number; from: string } | null>(null);
  const addrOf = (target: EventTarget | null) => {
    const td = (target as HTMLElement | null)?.closest?.('td') as HTMLTableCellElement | null;
    const tr = td?.parentElement as HTMLTableRowElement | undefined;
    if (!td || !tr?.dataset.er) return null;
    const ci = td.cellIndex - 2;
    if (ci < 0) return null;
    const col = ci < p1.length ? (/^c\d+$/.test(p1[ci].key) ? colName(Number(p1[ci].key.slice(1))) : colName(ci + 1)) : colName(p2Start + ci - p1.length);
    return `${col}${tr.dataset.er}`;
  };
  const editingFormula = (): HTMLInputElement | null => {
    const el = document.activeElement as HTMLInputElement | null;
    if (!el || el.tagName !== 'INPUT' || !isFormulaText(el.value)) return null;
    if (!(el.dataset.c || el.classList.contains('bv-fx-input'))) return null;
    const pos = el.selectionStart ?? el.value.length;
    return /[=+\-*/^(,;:<>&]\s*$/.test(el.value.slice(0, pos)) ? el : null;
  };
  // ===== Gợi ý hàm khi gõ công thức (SUM, SUMPRODUCT, SUBTOTAL…) =====
  type Hint = { el: HTMLInputElement; start: number; end: number; items: typeof FUNC_INFO; sel: number; sig: (typeof FUNC_INFO)[number] | null; x: number; y: number };
  const [hint, setHint] = useState<Hint | null>(null);
  // Đang gõ công thức: tô màu các ô / vùng được tham chiếu (như Excel) – chỉ tô, không sửa giá trị ô nào
  const REF_COLORS = ['#2563eb', '#dc2626', '#7c3aed', '#059669', '#d97706', '#db2777'];
  const clearRefs = () => gridRef.current?.querySelectorAll('.fref').forEach(e => { e.classList.remove('fref'); (e as HTMLElement).style.removeProperty('--fref'); });
  const highlightRefs = (text: string | null) => {
    clearRefs();
    const t = gridRef.current; if (!t || !text || !isFormulaText(text)) return;
    const body = text.replace(/\[[^\]]*\]/g, ' ');
    const re = /\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?/g;
    let m: RegExpExecArray | null; let n = 0;
    const ciOf = (L: string) => { const i1 = letterP1.get(L); if (i1 !== undefined) return 2 + i1; const k = colNum(L) - p2Start; return k >= 0 && k < p2keys.length ? 2 + p1.length + k : -1; };
    while ((m = re.exec(body))) {
      if (/[A-Za-z(]/.test(body[re.lastIndex] || '')) continue;  // tên hàm (SUM(…)
      const c1 = colNum(m[1].toUpperCase()), r1 = Number(m[2]);
      const c2 = m[3] ? colNum(m[3].toUpperCase()) : c1, r2 = m[4] ? Number(m[4]) : r1;
      const color = REF_COLORS[n++ % REF_COLORS.length];
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
        t.querySelectorAll<HTMLTableRowElement>(`tbody tr[data-er="${r}"]`).forEach(tr => {
          for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
            const ci = ciOf(colName(c)); const td = ci >= 0 ? tr.cells[ci] as HTMLElement | undefined : undefined;
            if (td) { td.classList.add('fref'); td.style.setProperty('--fref', color); }
          }
        });
      }
    }
  };
  // Chọn ô công thức: tô màu các ô trong công thức (sau khi xử lý rời ô trước)
  const showRefsSoon = (f: string | null) => { setTimeout(() => (f ? highlightRefs(f) : clearRefs()), 170); };
  const updateHint = (el: HTMLInputElement | null) => {
    highlightRefs(el?.value ?? null);
    if (!el || !isFormulaText(el.value)) { setHint(null); return; }
    const pos = el.selectionStart ?? el.value.length;
    const before = el.value.slice(0, pos);
    const rc = el.getBoundingClientRect();
    const m = /(^|[=+\-*/^(,;:<>&\s])([A-Za-z]{1,12})$/.exec(before);
    const items = m ? FUNC_INFO.filter(f => f.name.startsWith(m[2].toUpperCase()) && !/^[A-Z]{1,3}\d/.test(m[2])) : [];
    // hàm đang mở ngoặc gần nhất → hiện cú pháp
    let depth = 0, sig: Hint['sig'] = null;
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === ')') depth++;
      else if (ch === '(') { if (depth) depth--; else { const fm = /([A-Za-z]+)$/.exec(before.slice(0, i)); sig = fm ? FUNC_INFO.find(f => f.name === fm[1].toUpperCase()) || null : null; break; } }
    }
    if (!items.length && !sig) { setHint(null); return; }
    setHint({ el, start: m ? pos - m[2].length : pos, end: pos, items: items.slice(0, 8), sel: 0, sig, x: rc.left, y: rc.bottom + 2 });
  };
  const applyHint = (h: Hint, i: number) => {
    const f = h.items[i]; if (!f) return;
    const text = f.name === 'SUBTOTAL' ? 'SUBTOTAL(9;' : f.name + '(';
    putText(h.el, h.start, h.end, text);
    setTimeout(() => updateHint(h.el), 0);
  };
  hintKeyRef.current = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!hint || !hint.items.length || hint.el !== e.currentTarget) return false;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); setHint({ ...hint, sel: (hint.sel + (e.key === 'ArrowDown' ? 1 : hint.items.length - 1)) % hint.items.length }); return true; }
    if (e.key === 'Tab' || e.key === 'Enter') { e.preventDefault(); applyHint(hint, hint.sel); return true; }
    if (e.key === 'Escape') { e.preventDefault(); setHint(null); return true; }
    return false;
  };
  const putText = (el: HTMLInputElement, start: number, end: number, text: string) => {
    const v = el.value.slice(0, start) + text + el.value.slice(end);
    if (el.classList.contains('bv-fx-input')) setFxText(v); else el.value = v;
    el.value = v;
    if (!el.classList.contains('bv-fx-input')) setFxText(v);
    highlightRefs(v);
    requestAnimationFrame(() => { el.focus(); el.setSelectionRange(start + text.length, start + text.length); });
  };
  const onGridMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0) return;
    // Bấm vào phần trống của ô (viền / khoảng đệm quanh ô nhập, dòng cao) → vẫn chọn đúng ô đó
    const tgt = e.target as HTMLElement;
    if (tgt.tagName === 'TD' && !e.shiftKey && !e.ctrlKey && !e.metaKey && !editingFormula()) {
      const inp = tgt.querySelector<HTMLInputElement>('input[data-c]:not([type=checkbox])');
      if (inp && !inp.readOnly) { e.preventDefault(); setTimeout(() => inp.focus(), 0); }
    }
    const gtr = (e.target as HTMLElement).closest?.('tr.grp-row') as HTMLTableRowElement | null;
    if (gtr && !editingFormula()) {  // dòng nhóm / tổng: hiện công thức gốc của ô (vd =SUBTOTAL(9,I6:I30)) trên thanh fx
      const td = (e.target as HTMLElement).closest('td') as HTMLTableCellElement | null;
      const sr = (d?.structure?.[activeSheet] || []).find(x => String(x.sid) === gtr.dataset.sid);
      if (td && sr) {
        const i = td.cellIndex - 2; const a = addrOf(td);
        const ff = i >= 0 && i < p1.length ? sr.f?.[i] : null;
        setActive(null); setCellKind(ff ? 'f' : 'v'); showRefsSoon(ff || null);
        setFx({ ref: '', name: a ? 'Ô ' + a : 'Ô', expr: ff ? `${ff}  →  ${td.innerText.trim()}` : td.innerText.trim() || '(trống)' });
      }
    }
    if ((e.target as HTMLElement).classList?.contains('fill-h')) {
      e.preventDefault(); e.stopPropagation(); (document.activeElement as HTMLElement | null)?.blur?.();
      const g = range(); if (g) fillRef.current = { g, to: g.r2 };
      return;
    }
    const rnTd = (e.target as HTMLElement).closest?.('td.rn') as HTMLTableCellElement | null;
    const rnTr = rnTd?.parentElement as HTMLTableRowElement | undefined;
    if (rnTd && rnTr?.dataset.ri !== undefined && !editingFormula()) {  // bấm số Dòng → chọn cả dòng
      e.preventDefault(); (document.activeElement as HTMLElement | null)?.blur?.();
      const ri = Number(rnTr.dataset.ri);
      if (e.shiftKey && selA) { setSelA({ r: selA.r, c: 2 }); setSelB({ r: ri, c: lastCol() }); return; }
      setSelA({ r: ri, c: 2 }); setSelB({ r: ri, c: lastCol() }); lineDrag.current = 'row';
      return;
    }
    const fEl = editingFormula();
    if (fEl && e.target !== fEl) {
      const a = addrOf(e.target);
      if (a) {
        e.preventDefault();
        const pos = fEl.selectionStart ?? fEl.value.length;
        pointRef.current = { el: fEl, start: pos, end: pos + a.length, from: a };
        putText(fEl, pos, fEl.selectionEnd ?? pos, a);
        return;
      }
    }
    const p = cellPos(e.target); if (!p || p.c < 2) return;
    if (e.shiftKey && selA) { setSelB({ r: p.r, c: p.c }); e.preventDefault(); return; }
    if ((e.ctrlKey || e.metaKey) && range()) { const g0 = range()!; setExtraR(x => [...x, g0]); } else if (extraR.length) setExtraR([]);
    setSelA({ r: p.r, c: p.c }); setSelB({ r: p.r, c: p.c }); dragging.current = true;
    if (!p.td.querySelector('input:not([type=checkbox])')) {  // ô không sửa: thanh fx hiện giá trị của chính ô đó
      setActive(null);
      setCellKind(p.td.classList.contains('xf') || p.td.classList.contains('fx') || p.td.classList.contains('pctd') ? 'f' : p.td.querySelector('.cx-badge') ? 'a' : 'v');
      showRefsSoon(null);
      const a = addrOf(p.td);
      const k2 = p2keys[p.c - 2 - p1.length]; const it2 = rows[p.r];
      if (it2 && ['amt', 'diff', 'pct', 'k3_diff', 'k3_pct'].includes(k2) && builtinFormula(it2, k2)) { showFx(it2, k2); return; }
      const ff = it2 && p.c - 2 < p1.length ? rowByItem.get(it2.id)?.f?.[p.c - 2] : null;
      if (ff) { setCellKind('f'); showRefsSoon(ff); setFx({ ref: '', name: a ? 'Ô ' + a : 'Ô', expr: `${ff}  →  ${cellText(p.td)}` }); return; }
      setFx({ ref: '', name: a ? 'Ô ' + a : 'Ô', expr: cellText(p.td) || '(trống)' });
    }
  };
  const onGridMouseOver = (e: React.MouseEvent) => {
    const fl = fillRef.current;
    if (fl && (e.buttons & 1)) {
      const tr = (e.target as HTMLElement).closest?.('tr[data-ri]') as HTMLTableRowElement | null;
      if (!tr) return;
      fl.to = Math.max(fl.g.r2, Number(tr.dataset.ri));
      gridRef.current?.querySelectorAll('.cfill').forEach(x => x.classList.remove('cfill'));
      for (let r = fl.g.r2 + 1; r <= fl.to; r++) { const row = rowEl(r); for (let c = fl.g.c1; c <= fl.g.c2; c++) row?.cells[c]?.classList.add('cfill'); }
      return;
    }
    if (lineDrag.current === 'row' && (e.buttons & 1)) {
      const tr = (e.target as HTMLElement).closest?.('tr[data-ri]') as HTMLTableRowElement | null;
      if (tr) setSelB({ r: Number(tr.dataset.ri), c: lastCol() });
      return;
    }
    const pt = pointRef.current;
    if (pt && (e.buttons & 1)) {  // kéo khi đang trỏ ô → vùng M7:M10
      const a = addrOf(e.target); if (!a) return;
      const text = a === pt.from ? a : `${pt.from}:${a}`;
      putText(pt.el, pt.start, pt.end, text); pt.end = pt.start + text.length;
      return;
    }
    if (!dragging.current || !(e.buttons & 1)) return;
    const p = cellPos(e.target); if (!p || p.c < 2) return;
    if (!selB || selB.r !== p.r || selB.c !== p.c) { setSelB({ r: p.r, c: p.c }); window.getSelection()?.removeAllRanges(); }
  };
  const cellText = (td: HTMLTableCellElement) => { const inp = td.querySelector('input'); return (inp ? inp.value : td.innerText).replace(/[\t\n]+/g, ' ').trim(); };
  // Thống kê vùng chọn như thanh trạng thái Excel (SUM / AVERAGE / COUNT / MIN / MAX)
  const [selStats, setSelStats] = useState<{ sum: number; avg: number; count: number; min: number; max: number } | null>(null);
  useEffect(() => {
    const g = range();
    if (!g || (g.r1 === g.r2 && g.c1 === g.c2)) { setSelStats(null); return; }
    const vals: number[] = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const tr = rowEl(r); if (!tr) continue;
      for (let c = g.c1; c <= g.c2; c++) {
        const td = tr.cells[c]; if (!td) continue;
        const t = cellText(td).replace(/[▲▼?]/g, ' ').trim().split(/\s+/)[0] || '';
        if (!t || t.endsWith('%')) continue;
        const v = parseVn(t); if (v !== null && /\d/.test(t)) vals.push(v);
      }
    }
    setSelStats(vals.length ? { sum: vals.reduce((a, b) => a + b, 0), avg: vals.reduce((a, b) => a + b, 0) / vals.length, count: vals.length, min: Math.min(...vals), max: Math.max(...vals) } : null);
  }, [selA, selB, edits]);
  const multi = () => { const g = range(); return !!g && (g.r1 !== g.r2 || g.c1 !== g.c2); };
  // Bộ nhớ copy nội bộ: giữ công thức + địa chỉ nguồn để dán / kéo xuống dịch tham chiếu tương đối như Excel
  const clipRef = useRef<{ text: string; cells: { formula: string | null; text: string; col: number; row: number | null }[][]; rowsData?: Edit[] } | null>(null);
  const srcCell = (r: number, c: number) => {
    const tr = rowEl(r); const td = tr?.cells[c]; const it = rows[r];
    const inp = td?.querySelector('input:not([type=checkbox])') as HTMLInputElement | null;
    const field = inp?.dataset.c;
    // Ô có công thức (người gõ, hoặc công thức sẵn như ĐG tổng =ROUND(M7*O7+N7;-2)) → copy công thức như Excel
    const pk = c - 2 - p1.length >= 0 ? p2keys[c - 2 - p1.length] : null;
    const fileF = it && c - 2 >= 0 && c - 2 < p1.length ? rowByItem.get(it.id)?.f?.[c - 2] || null : null;  // công thức sẵn trong file NTP (vd =ROUND(H7*$D7,2))
    const formula = (it && field ? formulaOfCell(it, field) : null) || fileF || (it && pk ? builtinFormula(it, pk)?.formula || null : null);
    const a = td ? addrOf(td) : null;
    const m = a ? /^([A-Z]+)(\d+)$/.exec(a) : null;
    return { formula, text: td ? cellText(td) : '', col: m ? colNum(m[1]) : c, row: m ? Number(m[2]) : null };
  };
  const shiftFormula = (f: string, dr: number, dc: number) => f.replace(/\[[^\]]*\]|(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])/g, (m, cAbs, col, rAbs, row) => {
    if (m.startsWith('[')) return m;
    const nc = cAbs ? col : colName(Math.max(1, colNum(col) + dc));
    const nr = rAbs ? row : String(Math.max(1, Number(row) + dr));
    return `${cAbs}${nc}${rAbs}${nr}`;
  });
  // Ghi một ô từ nguồn copy: công thức → dịch theo khoảng cách dòng/cột Excel; giá trị → như dán thường
  const putFromSource = (r: number, c: number, src: { formula: string | null; text: string; col: number; row: number | null }) => {
    const tr = rowEl(r); const td = tr?.cells[c]; const it = rows[r];
    const inp = td?.querySelector('input:not([type=checkbox])') as HTMLInputElement | null;
    if (!it || !inp || inp.readOnly || !inp.dataset.c) return 0;
    if (src.formula) {
      const a = addrOf(td!); const m = a ? /^([A-Z]+)(\d+)$/.exec(a) : null;
      const dr = m && src.row ? Number(m[2]) - src.row : 0, dc = m ? colNum(m[1]) - src.col : 0;
      const nf = shiftFormula(src.formula, dr, dc);
      // Trùng đúng công thức sẵn của ô đích (vd dán =ROUND(M8*O8+N8;-2) vào P8) → giữ công thức sẵn, bỏ số gõ đè
      const pk = c - 2 - p1.length >= 0 ? p2keys[c - 2 - p1.length] : null;
      if (pk === 'total' && inp.dataset.c === 'rv_total' && builtinFormulaPlain(it, pk) === nf) { editRow(it, { rv_total: null, f_rv_total: null }); return 1; }
      return commitCell(it, inp.dataset.c, nf, inp.dataset.c.endsWith('dep_pct') || inp.dataset.c === 'x_k3_dep') ? 1 : 0;
    }
    const pr = parseInto(inp.dataset.c, src.text);
    if (!pr.ok) return 0;
    editRow(it, { [inp.dataset.c]: pr.v, ...(formulaOfCell(it, inp.dataset.c) ? { ['f_' + inp.dataset.c]: null } : {}) });
    return 1;
  };
  const fillDown = (g = range()) => {  // Ctrl+D / kéo ô vuông góc: chép dòng đầu (hoặc nhóm dòng nguồn) xuống
    if (!g) return 0;
    let n = 0;
    for (let c = g.c1; c <= g.c2; c++) { const src = srcCell(g.r1, c); for (let r = g.r1 + 1; r <= g.r2; r++) n += putFromSource(r, c, src); }
    return n;
  };
  const copySel = async () => {
    const g = range(); if (!g) return;
    const cells: { formula: string | null; text: string; col: number; row: number | null }[][] = [];
    for (let r = g.r1; r <= g.r2; r++) { const row: typeof cells[number] = []; for (let c = g.c1; c <= g.c2; c++) row.push(srcCell(r, c)); cells.push(row); }
    const lines: string[] = [];
    for (let r = g.r1; r <= g.r2; r++) { const tr = rowEl(r); if (!tr) continue; const vals: string[] = []; for (let c = g.c1; c <= g.c2; c++) vals.push(tr.cells[c] ? cellText(tr.cells[c]) : ''); lines.push(vals.join('\t')); }
    const txt = lines.join('\n');
    // Copy nguyên dòng (chọn cả dòng) → nhớ số liệu dòng để "Chèn dòng đã copy"
    const full = g.c1 <= 2 && g.c2 >= lastCol();
    const rowsData = full ? rows.slice(g.r1, g.r2 + 1).map(it => rowPatchOf(it)) : undefined;
    clipRef.current = { text: txt, cells, rowsData };
    try { await navigator.clipboard.writeText(txt); } catch {
      const ta = document.createElement('textarea'); ta.value = txt; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove();
    }
    toast(`Đã copy ${g.r2 - g.r1 + 1} dòng × ${g.c2 - g.c1 + 1} cột`);
  };
  // Giá trị dán / xóa vào một ô sửa được
  const parseInto = (field: string, raw: string): { ok: boolean; v: any } => {
    const t = raw.trim();
    if (t === '') return { ok: true, v: null };
    if (field === 'rv_note' || field === 'x_k3_note' || field.startsWith('add_') || (field.startsWith('x_') && !K3_NUMF.has(field) && xcols.find(x => 'x_' + x.id === field)?.type !== 'number')) return { ok: true, v: t };
    const isPct = field.endsWith('dep_pct') || field === 'x_k3_dep';
    let v = parseVn(isPct ? t.replace('%', '').trim() : t);
    if (v === null || v < 0) return { ok: false, v: null };
    if (isPct && (v > 1 || t.includes('%'))) v = v / 100;
    if (PRICE_FIELDS.has(field)) v = roundUnit(v, ru);
    return { ok: true, v };
  };
  const writeCells = (startR: number, startC: number, grid: string[][]) => {
    const patches = new Map<number, Edit>(); let bad = 0, n = 0;
    grid.forEach((vals, dr) => {
      const tr = rowEl(startR + dr); const it = rows[startR + dr]; if (!tr || !it) return;
      vals.forEach((raw, dc) => {
        const inp = tr.cells[startC + dc]?.querySelector('input:not([type=checkbox])') as HTMLInputElement | null;
        if (!inp || inp.readOnly || !inp.dataset.c) return;
        const r = parseInto(inp.dataset.c, raw);
        if (!r.ok) { bad++; return; }
        patches.set(it.id, { ...(patches.get(it.id) || {}), [inp.dataset.c]: r.v }); n++;
      });
    });
    patches.forEach((pt, iid) => { const it = itemById.get(iid); if (it) editRow(it, pt); });
    return { n, bad };
  };
  const onGridPaste = (e: React.ClipboardEvent) => {
    const txt = e.clipboardData.getData('text/plain'); if (!txt) return;
    if (doPaste(txt, e.target, false)) e.preventDefault();
  };
  // Dán (Ctrl+V hoặc nút Dán): trả true nếu đã xử lý
  const doPaste = (txt: string, target: EventTarget | null, force: boolean): boolean => {
    // Dán từ copy trong bảng (có công thức) → dán công thức, dịch tham chiếu như Excel
    const clip = clipRef.current;
    const ae = document.activeElement as HTMLInputElement | null;
    const editingText = ae?.tagName === 'INPUT' && ae.value !== (ae.dataset.orig ?? ae.value);
    if (clip && clip.text.replace(/\r/g, '') === txt.replace(/\r/g, '').replace(/\n$/, '') && clip.cells.some(r => r.some(x => x.formula)) && !editingText) {
      const g = range(); const st = g ? { r: g.r1, c: g.c1 } : cellPos(target); if (!st) return true;
      const h = clip.cells.length, w = clip.cells[0].length;
      const rr = g && (g.r2 - g.r1 + 1) > h ? g.r2 - g.r1 + 1 : h, cc = g && (g.c2 - g.c1 + 1) > w ? g.c2 - g.c1 + 1 : w;
      let n = 0;
      for (let i = 0; i < rr; i++) for (let j = 0; j < cc; j++) n += putFromSource(st.r + i, st.c + j, clip.cells[i % h][j % w]);
      (document.activeElement as HTMLElement | null)?.blur?.();
      toast(n ? `Đã dán ${n} ô (kèm công thức) – bấm Lưu để ghi` : 'Vùng dán không có ô sửa được', n ? 'ok' : 'err');
      return true;
    }
    let grid = txt.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
    const single = grid.length === 1 && grid[0].length === 1;
    const g = range();
    if (single && !multi() && !force) return false;  // dán 1 giá trị vào ô đang gõ: để trình duyệt xử lý
    let start = g ? { r: g.r1, c: g.c1 } : cellPos(target);
    if (!g && !start) return true;
    if (single && g) grid = Array.from({ length: g.r2 - g.r1 + 1 }, () => Array(g.c2 - g.c1 + 1).fill(grid[0][0]));  // 1 giá trị → điền cả vùng
    if (!g) { const p = cellPos(target); if (p) start = { r: p.r, c: p.c }; }
    (document.activeElement as HTMLElement | null)?.blur?.();
    const res = writeCells(start!.r, start!.c, grid);
    toast(res.n ? `Đã dán ${res.n} ô – bấm Lưu để ghi${res.bad ? ` (bỏ qua ${res.bad} ô không phải số)` : ''}` : 'Vùng dán không có ô sửa được (ô vàng)', res.n ? 'ok' : 'err');
    return true;
  };
  // ===== Định dạng ô (thanh công cụ) =====
  const colIdOf = (ci: number) => (ci - 2 < p1.length ? `p1:${p1[ci - 2]?.key}` : `p2:${p2keys[ci - 2 - p1.length]}`);
  const stylesOf = (it: Item): Record<string, CellStyle> => ((edits[it.id] as any)?.st !== undefined ? (edits[it.id] as any).st : it.st) || {};
  const anchorStyle = (): CellStyle => { const g = range(); const it = g ? rows[g.r1] : undefined; return it && g ? stylesOf(it)[colIdOf(g.c1)] || {} : {}; };
  const [colorPop, setColorPop] = useState<'c' | 'bg' | null>(null);
  useEffect(() => { if (!colorPop) return; const h = (e: MouseEvent) => { if (!(e.target as HTMLElement).closest('.xl-pal-wrap')) setColorPop(null); }; window.addEventListener('mousedown', h); return () => window.removeEventListener('mousedown', h); }, [colorPop]);
  const applyStyle = (fn: (s: CellStyle) => CellStyle) => {
    const rs = allRanges(); if (!rs.length) { toast('Chọn ô / vùng ô trước', 'err'); return; }
    const acc = new Map<number, { it: Item; cur: Record<string, CellStyle> }>();
    for (const g of rs) for (let r = g.r1; r <= g.r2; r++) {
      const it = rows[r]; if (!it) continue;
      const a = acc.get(it.id) || { it, cur: { ...stylesOf(it) } };
      for (let c = g.c1; c <= g.c2; c++) {
        const id = colIdOf(c);
        const ns = fn({ ...(a.cur[id] || {}) });
        (Object.keys(ns) as (keyof CellStyle)[]).forEach(k => { if (ns[k] === undefined || ns[k] === false || ns[k] === '') delete ns[k]; });
        if (Object.keys(ns).length) a.cur[id] = ns; else delete a.cur[id];
      }
      acc.set(it.id, a);
    }
    acc.forEach(({ it, cur }) => editRow(it, { st: Object.keys(cur).length ? cur : null }));
  };
  const toolbarPaste = async (valuesOnly = false) => {
    let t = '';
    try { t = await navigator.clipboard.readText(); } catch { t = clipRef.current?.text || ''; }
    if (!t) { toast('Chưa có gì để dán – Ctrl+C trước', 'err'); return; }
    if (valuesOnly) { const keep = clipRef.current; clipRef.current = null; doPaste(t, null, true); clipRef.current = keep; }
    else doPaste(t, null, true);
  };
  // Chuột phải vào ô: nếu ô nằm ngoài vùng chọn thì chọn ô đó (như Excel) rồi mở menu
  const ctxRef = useRef<(e: React.MouseEvent, sid: number, it?: Item) => void>(() => {});
  ctxRef.current = (e, sid, it) => {
    e.preventDefault();
    const pos = cellPos(e.target); const g = range();
    if (pos && pos.c >= 2 && !(g && pos.r >= g.r1 && pos.r <= g.r2 && pos.c >= g.c1 && pos.c <= g.c2)) { setSelA({ r: pos.r, c: pos.c }); setSelB({ r: pos.r, c: pos.c }); }
    if (it) setCur(it.id);
    setCurSid(sid);
    setMenu({ x: e.clientX, y: e.clientY, sid, it });
  };
  gridKeysRef.current = (e: KeyboardEvent) => {
    if (!range() || tab !== 'detail') return false;
    const k = e.key.toLowerCase();
    if ((e.ctrlKey || e.metaKey) && k === 'c') { e.preventDefault(); copySel(); return true; }
    if ((e.ctrlKey || e.metaKey) && k === 'v') { e.preventDefault(); toolbarPaste(); return true; }
    if ((e.ctrlKey || e.metaKey) && k === 'd') { e.preventDefault(); const n = fillDown(); toast(n ? `Đã điền xuống ${n} ô – bấm Lưu để ghi` : 'Không có ô sửa được để điền'); return true; }
    if (e.key === 'Insert') { e.preventDefault(); if (clipRef.current?.rowsData?.length) insertCopiedRows(); else addRow('after'); return true; }
    if (e.key === 'Delete') { e.preventDefault(); let n = 0; for (const g of allRanges()) n += writeCells(g.r1, g.c1, Array.from({ length: g.r2 - g.r1 + 1 }, () => Array(g.c2 - g.c1 + 1).fill(''))).n; if (n) toast(`Đã xóa ${n} ô – bấm Lưu để ghi`); return true; }
    return false;
  };
  const onGridKeyDown = (e: React.KeyboardEvent) => {
    const ae = document.activeElement as HTMLInputElement | null;
    const k = e.key.toLowerCase();
    if ((e.ctrlKey || e.metaKey) && (k === 'z' || k === 'y') && !(ae?.tagName === 'INPUT' && ae.value !== (ae.dataset.orig ?? ae.value))) { e.preventDefault(); k === 'z' ? undo() : redo(); return; }
    if ((e.ctrlKey || e.metaKey) && (k === 'b' || k === 'i' || k === 'u') && range()) { e.preventDefault(); const kk = k as 'b' | 'i' | 'u'; const on = !anchorStyle()[kk]; applyStyle(s => ({ ...s, [kk]: on })); return; }
    const partialText = ae?.tagName === 'INPUT' && !ae.classList.contains('rdy') && ae.selectionStart !== null && (ae.selectionStart !== 0 || ae.selectionEnd !== ae.value.length);
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c' && range() && (multi() || !partialText)) { e.preventDefault(); copySel(); }
    else if (e.key === 'Insert' || ((e.ctrlKey || e.metaKey) && (e.key === '+' || e.key === '='))) {  // chèn: mặc định xuống dưới dòng đang chọn
      e.preventDefault(); if (clipRef.current?.rowsData?.length) insertCopiedRows(); else addRow('after');
    }
    else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd' && multi()) { e.preventDefault(); const n = fillDown(); toast(n ? `Đã điền xuống ${n} ô – bấm Lưu để ghi` : 'Không có ô sửa được để điền'); }
    else if ((e.key === 'Delete' || (e.key === 'Backspace' && (multi() || allRanges().length > 1))) && (multi() || allRanges().length > 1 || !(ae?.tagName === 'INPUT') || ae.classList.contains('rdy'))) {  // xóa cả vùng chọn (kể cả nhiều vùng Ctrl)
      e.preventDefault(); let n = 0;
      for (const g of allRanges()) n += writeCells(g.r1, g.c1, Array.from({ length: g.r2 - g.r1 + 1 }, () => Array(g.c2 - g.c1 + 1).fill(''))).n;
      if (n) toast(`Đã xóa ${n} ô – bấm Lưu để ghi`);
    } else if (e.key === 'Escape') { setSelA(null); setSelB(null); setExtraR([]); }
  };
  // Công thức Excel của các ô tool tự tính (giống công thức trong file xuất): ĐG tổng, Thành tiền, Chênh lệch, % chênh lệch
  const builtinFormula = (it: Item, key: string): { addr: string; formula: string; value: string } | null => {
    const c = calcs.get(it.id); const r = rowByItem.get(it.id);
    if (!c || !r || it.is_added) return null;
    const L = (k: string) => { const i = p2keys.indexOf(k); return i < 0 ? null : `${colName(p2Start + i)}${r.row}`; };
    const mp = cfg?.mapping || {};
    const ntpAmt = mp.amount ? `${colName(mp.amount)}${r.row}` : mp.quantity && mp.price ? `${colName(mp.quantity)}${r.row}*${colName(mp.price)}${r.row}` : null;
    const rnd = (e: string) => (ru ? `ROUND(${e};${-Math.round(Math.log10(ru))})` : e);
    const addr = L(key === 'total' ? 'total' : key) || '';
    if (key === 'total') {
      if (!c.hasParts || (c.rv.rv_total !== null && c.rv.rv_total !== undefined) || !L('vt') || !L('nc')) return null;
      const vt = c.dep !== null && L('dep') ? `${L('vt')}*${L('dep')}` : L('vt')!;
      return { addr, formula: `=${rnd(`${vt}+${L('nc')}${c.rv.rv_machine ? `+${c.rv.rv_machine}` : ''}`)}`, value: money(c.total) || '0' };
    }
    if (key === 'amt' && L('total') && L('qty')) return { addr, formula: `=${L('total')}*${L('qty')}`, value: money(c.amt) || '0' };
    if (key === 'diff' && L('amt') && ntpAmt) return { addr, formula: `=${L('amt')}-${ntpAmt}`, value: money(c.diff) || '0' };
    if (key === 'pct' && L('diff') && ntpAmt) return { addr, formula: `=${L('diff')}/${ntpAmt}`, value: c.ntpAmt ? `${(c.diff / c.ntpAmt * 100).toFixed(1)}%` : '' };
    if (key.startsWith('k3_')) {  // Phần 3: ô chưa sửa = tham chiếu ô Phần 2 cùng dòng
      const own = (k: string) => { const f = K3_FIELD[k]; const v = (edits[it.id] as any)?.[f] !== undefined ? (edits[it.id] as any)[f] : (it as any)[f]; return typeof v === 'number'; };
      const base = key.slice(3);
      if (['k3_qty', 'k3_vt', 'k3_nc', 'k3_dep'].includes(key)) return own(key) || !L(base) ? null : { addr, formula: `=${L(base)}`, value: String(keyValue(it, key) ?? '') };
      if (key === 'k3_total') {
        if (own('k3_total')) return null;
        const touched = ['k3_vt', 'k3_nc', 'k3_dep'].some(own);
        if (c.k3.parts && ((c.rv.rv_total === null || c.rv.rv_total === undefined) || touched) && L('k3_vt') && L('k3_nc')) {
          const vt = c.k3.dep !== null && L('k3_dep') ? `${L('k3_vt')}*${L('k3_dep')}` : L('k3_vt')!;
          return { addr, formula: `=${rnd(`${vt}+${L('k3_nc')}${c.rv.rv_machine ? `+${c.rv.rv_machine}` : ''}`)}`, value: money(c.k3.total) || '0' };
        }
        return L('total') ? { addr, formula: `=${L('total')}`, value: money(c.k3.total) || '0' } : null;
      }
      if (key === 'k3_amt' && !own('k3_amt') && L('k3_total') && L('k3_qty')) return { addr, formula: `=${L('k3_total')}*${L('k3_qty')}`, value: money(c.k3.amt) || '0' };
      if (key === 'k3_diff' && L('k3_amt') && ntpAmt) return { addr, formula: `=${L('k3_amt')}-${ntpAmt}`, value: money(c.k3.diff) || '0' };
      if (key === 'k3_pct' && L('k3_diff') && ntpAmt) return { addr, formula: `=${L('k3_diff')}/${ntpAmt}`, value: c.ntpAmt ? `${(c.k3.diff / c.ntpAmt * 100).toFixed(1)}%` : '' };
    }
    return null;
  };
  // Công thức sẵn của ô khi chưa gõ đè (dùng khi dán công thức ĐG tổng trở lại)
  const builtinFormulaPlain = (it: Item, key: string) => {
    const c = calcs.get(it.id); const r = rowByItem.get(it.id);
    if (!c || !r || it.is_added || !c.hasParts || key !== 'total') return null;
    const L = (k: string) => { const i = p2keys.indexOf(k); return i < 0 ? null : `${colName(p2Start + i)}${r.row}`; };
    if (!L('vt') || !L('nc')) return null;
    const vt = c.dep !== null && L('dep') ? `${L('vt')}*${L('dep')}` : L('vt')!;
    const e = `${vt}+${L('nc')}${c.rv.rv_machine ? `+${c.rv.rv_machine}` : ''}`;
    return `=${ru ? `ROUND(${e};${-Math.round(Math.log10(ru))})` : e}`;
  };
  const builtinRef = useRef<string | null>(null);
  const showFx = (it: Item, key: string) => {
    const b = builtinFormula(it, key);
    if (b) { setCellKind('f'); showRefsSoon(b.formula); setFx({ ref: `${it.boq_stt || ''} ${it.boq_description || ''}`.trim(), name: 'Ô ' + b.addr, expr: `${b.formula}  →  ${b.value}` }); return; }
    const f = formulaOf(it, calcs.get(it.id)!, key);
    setFx(f ? { ref: `${it.boq_stt || ''} ${it.boq_description || ''}`.trim(), ...f } : null);
  };
  // ===== Công thức Excel ngay trên bảng: =M7*1,1  =SUM(M7:M12)  =[ĐG VT NTP]*0,9 =====
  // Địa chỉ ô: chữ cột (hàng A B C… trên tiêu đề) + số Dòng (cột Dòng = dòng trong file Excel).
  const sheetRowsByExcel = useMemo(() => {
    const m = new Map<number, SRow>();
    for (const r of d?.structure?.[activeSheet] || []) { const it = r.item_id ? itemById.get(r.item_id) : undefined; if (!it?.is_added && !m.has(r.row)) m.set(r.row, r); }
    return m;
  }, [d, activeSheet, itemById]);
  const letterP1 = useMemo(() => new Map(p1.map((c, i) => [/^c\d+$/.test(c.key) ? colName(Number(c.key.slice(1))) : colName(i + 1), i])), [p1]);
  const keyValue = (it: Item, k: string): number | null => {
    const c = calcs.get(it.id); if (!c) return null;
    switch (k) {
      case 'qty': return c.qty; case 'vt': return c.rv.rv_material; case 'nc': return c.rv.rv_labor; case 'dep': return c.dep;
      case 'total': return c.total; case 'amt': return c.amt; case 'diff': return c.diff; case 'pct': return c.ntpAmt ? c.diff / c.ntpAmt : null; case 'lib': return it.lib_price;
      case 'k3_qty': return c.k3.qty; case 'k3_vt': return c.k3.vt; case 'k3_nc': return c.k3.nc; case 'k3_dep': return c.k3.dep;
      case 'k3_total': return c.k3.total; case 'k3_amt': return c.k3.amt; case 'k3_diff': return c.k3.diff; case 'k3_pct': return c.ntpAmt ? c.k3.diff / c.ntpAmt : null;
      default: { const v = xvals.vals.get(it.id)?.[k]; if (typeof v === 'number') return Number.isNaN(v) ? null : v;
        const e = (edits[it.id] as any)?.[k]; const w = e !== undefined ? e : (it as any)[k]; return typeof w === 'number' ? w : null; }
    }
  };
  const cellEnv = (col: string, row: number): number | null => {
    const r = sheetRowsByExcel.get(row);
    if (!r) return null;
    const i1 = letterP1.get(col);
    if (i1 !== undefined) {
      const it = r.item_id ? itemById.get(r.item_id) : undefined;
      const nf = NTP_EDIT[p1[i1].field || ''];
      if (it && nf) return calcs.get(it.id)?.ntp[NTP_VAL[nf]] ?? null;
      if (it && p1[i1].field === 'amount') return calcs.get(it.id)?.ntpAmt ?? null;
      const v = structured ? r.cells[i1] : (it ? (p1[i1].get(r, it) as any) : null);
      return typeof v === 'number' ? v : v == null || v === '' ? null : parseVn(String(v).replace('%', '')) ;
    }
    const k = p2keys[colNum(col) - p2Start];
    const it = r.item_id ? itemById.get(r.item_id) : undefined;
    return k && it ? keyValue(it, k) : null;
  };
  const evalCell = (it: Item, raw: string) =>
    evalFormula(raw, rowEnv(it, calcs.get(it.id)!, k => xvals.vals.get(it.id)?.[k] ?? null), cellEnv);
  const formulaOfCell = (it: Item, field: string): string | null => {
    const e = (edits[it.id] as any)?.['f_' + field];
    return e !== undefined ? e : (it as any)['f_' + field] ?? null;
  };
  // Tự tính lại các ô có công thức khi số liệu thay đổi (như Excel)
  const recalcPasses = useRef(0);
  useEffect(() => {
    if (viewVer) return;
    const patches: [Item, Edit][] = [];
    for (const it of items) {
      const fields = new Set([...Object.keys(it), ...Object.keys(edits[it.id] || {})].filter(k => k.startsWith('f_')).map(k => k.slice(2)));
      const patch: Edit = {};
      for (const f of fields) {
        const fm = formulaOfCell(it, f); if (!fm) continue;
        let v: number | null;
        try { v = evalCell(it, fm); } catch { continue; }
        if (v !== null && (f === 'rv_dep_pct' || f === 'boq_dep_pct' || f === 'x_k3_dep') && v > 1) v = v / 100;
        const cur = (edits[it.id] as any)?.[f] !== undefined ? (edits[it.id] as any)[f] : (it as any)[f];
        if (v !== null && (typeof cur !== 'number' || Math.abs(cur - v) > 1e-6)) (patch as any)[f] = v;
      }
      if (Object.keys(patch).length) patches.push([it, patch]);
    }
    if (!patches.length) { recalcPasses.current = 0; return; }
    if (++recalcPasses.current > 20) { toast('Công thức tham chiếu vòng – dừng tính lại', 'err'); return; }
    patches.forEach(([it, pt]) => editRow(it, pt));
  }, [calcs, xvals]);
  // Ghi giá trị / công thức vào ô (dùng chung cho ô trên bảng và thanh fx). Trả false nếu lỗi.
  const commitCell = (it: Item, field: string, raw: string, pct?: boolean): boolean => {
    const fm = formulaOfCell(it, field);
    // Gõ ĐG tổng: chỉ đổi ô ĐG tổng (ĐG VT / ĐG NC giữ nguyên như Excel); xóa trắng ô → về công thức VT × %KH + NC
    raw = raw.trim();
    if (raw === '') { editRow(it, { [field]: null, ...(fm ? { ['f_' + field]: null } : {}) }); return true; }
    if (isFormulaText(raw)) {
      raw = toFormula(raw);  // bắt đầu bằng + (hoặc -A1…) như Excel → lưu thành =…
      try {
        const v = evalCell(it, raw);
        if (v === null || v < 0) throw new Error('Kết quả phải là số ≥ 0');
        editRow(it, { [field]: pct && v > 1 ? v / 100 : PRICE_FIELDS.has(field) ? roundUnit(v, ru) : v, ['f_' + field]: raw });
        return true;
      } catch (er: any) { toast(`Công thức lỗi: ${er.message}`, 'err'); return false; }
    }
    if (field === 'rv_note' || field === 'x_k3_note' || field.startsWith('add_') || xById.get(field)?.type === 'text') { editRow(it, { [field]: raw }); return true; }
    let v = parseVn(pct ? raw.replace('%', '').trim() : raw);
    if (v === null || v < 0) { toast(pct ? 'Nhập tỷ lệ, ví dụ 20% hoặc 0,2' : 'Nhập số ≥ 0, ví dụ 1.250.000', 'err'); return false; }
    if (pct && (v > 1 || raw.includes('%'))) v = v / 100;
    if (PRICE_FIELDS.has(field)) v = roundUnit(v, ru);  // gõ số lẻ → tự làm tròn đơn giá
    if ((field === 'rv_quantity' || field === 'x_k3_qty') && ru) v = roundUnit(v, 1);  // KL làm tròn số nguyên
    editRow(it, { [field]: v, ...(fm ? { ['f_' + field]: null } : {}) });
    return true;
  };
  const commitRef = useRef(commitCell); commitRef.current = commitCell;
  const rowByItem = useMemo(() => new Map((d?.structure?.[activeSheet] || []).filter(x => x.item_id).map(x => [x.item_id!, x])), [d, activeSheet]);
  const numCell = (it: Item, idx: number, field: string, value: number | null, opts: { readOnly?: boolean; title?: string; qty?: boolean; pct?: boolean; addr?: string; ntp?: boolean; cls?: string; colId?: string; fileF?: string | null } = {}) => {
    const nf = stylesOf(it)[opts.colId || `p2:${FIELD_KEY[field] || field}`]?.nf;
    const shown = value === null || value === undefined ? '' : nfText(value, nf) ?? (opts.pct ? pctText(value) : opts.qty ? num(value) : money(value));
    const changedFromNtp = (edits[it.id] as any)?.[field] !== undefined;
    const fm = opts.readOnly ? null : formulaOfCell(it, field);
    const r = rowByItem.get(it.id);
    const cellF = () => fm || (BUILTIN_OF[field] ? builtinFormula(it, BUILTIN_OF[field])?.formula : opts.fileF) || null;
    const addr = opts.addr ?? (r && !it.is_added ? `${colName(p2Start + p2keys.indexOf(FIELD_KEY[field] || field))}${r.row}` : '');
    return (
      <td className={`r ${opts.ntp ? 'ntpedit' : 'edit'} ${opts.cls || ''} ${opts.readOnly ? 'ro' : ''} ${changedFromNtp ? 'dirty' : ''} ${fm ? 'hasf' : ''}`} title={fm ? `${addr} ${fm}` : opts.title || addr}>
        <input data-r={idx} data-c={field} defaultValue={shown} key={`${field}${it.id}-${shown}-${fm || ''}`} readOnly={opts.readOnly}
          tabIndex={opts.readOnly ? -1 : 0}
          onFocus={e => {
            // Như Excel: ô hiện giá trị, thanh fx hiện công thức của ĐÚNG ô này; gõ là thay nội dung, F2 / nhấp đúp để sửa công thức trong ô
            const bf = !fm && BUILTIN_OF[field] ? builtinFormula(it, BUILTIN_OF[field]) : !fm && opts.fileF ? { addr, formula: opts.fileF, value: shown } : null;
            builtinRef.current = bf ? bf.formula : null;
            if (!opts.readOnly) { setActive({ it, field, pct: opts.pct, addr }); setFxText(fm || bf?.formula || shown); }
            else setActive(null);
            setCellKind(fm || bf ? 'f' : 'v');
            showRefsSoon(fm || bf?.formula || null);
            if (bf) setFx({ ref: `${addr} · ${it.boq_stt || ''} ${it.boq_description || ''}`.trim(), name: 'Ô ' + addr, expr: `${bf.formula}  →  ${shown}` });
            else if (fm) setFx({ ref: `${addr} · ${it.boq_stt || ''} ${it.boq_description || ''}`.trim(), name: 'Ô ' + addr, expr: `${fm}  →  ${shown}` });
            else showFx(it, field === 'rv_total' ? 'total' : field);
            // Chọn ô: ô công thức hiện (và bôi chọn) CÔNG THỨC; ô giá trị hiện GIÁ TRỊ. Rời ô sẽ hiện lại kết quả.
            // Như Excel: ô vẫn hiện GIÁ TRỊ, công thức chỉ hiện trên thanh fx (F2 / nhấp đúp để sửa trong ô)
            e.currentTarget.value = shown;
            e.currentTarget.dataset.orig = shown;
            e.currentTarget.classList.add('rdy');  // chế độ "chọn ô" như Excel: không hiện con trỏ / bôi chữ
            e.currentTarget.select();
          }}
          onMouseDown={e => {  // bấm lại vào ô đang chọn: vẫn là chọn cả ô (nhấp đúp mới vào sửa)
            const el = e.currentTarget;
            if (e.detail < 2 && el.classList.contains('rdy') && document.activeElement === el) { e.preventDefault(); el.select(); requestAnimationFrame(() => { if (el.classList.contains('rdy')) el.select(); }); }
          }}
          onDoubleClick={e => { if (!opts.readOnly) enterEdit(e.currentTarget, cellF()); }}
          onKeyDown={e => {
            const el = e.currentTarget; const rdy = el.classList.contains('rdy');
            if (e.key === 'F2' && !opts.readOnly) { e.preventDefault(); enterEdit(el, cellF()); return; }
            if (e.key === 'Escape' && !rdy) { e.preventDefault(); el.value = shown; el.dataset.orig = shown; el.classList.add('rdy'); unwidenEdit(el); el.select(); return; }  // hủy sửa, về chọn ô
            if (rdy && (e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.key === 'Tab')) {  // chọn ô: mũi tên trái / phải sang ô bên cạnh
              const ins = Array.from(el.closest('tr')?.querySelectorAll<HTMLInputElement>('input[data-c]:not([readonly])') || []);
              const nx = ins[ins.indexOf(el) + (e.key === 'ArrowLeft' || (e.key === 'Tab' && e.shiftKey) ? -1 : 1)];
              if (nx) { e.preventDefault(); nx.focus(); }
              return;
            }
            if (rdy && !opts.readOnly && !e.ctrlKey && !e.metaKey && !e.altKey && (e.key.length === 1 || e.key === 'Backspace')) el.classList.remove('rdy');  // gõ → thay nội dung ô
            onCellKey(e);
          }}
          onBlur={e => {
            unwidenEdit(e.currentTarget);
            e.currentTarget.classList.remove('rdy');
            if (opts.readOnly) return;
            const raw = e.currentTarget.value.trim();
            const orig = e.currentTarget.dataset.orig ?? shown;
            e.currentTarget.classList.remove('fshow');
            if (raw === orig || (fm && raw === fm) || raw === shown) { e.currentTarget.value = shown; return; }  // không đổi → giữ nguyên công thức
            if (!commitRef.current(it, field, raw, opts.pct)) e.currentTarget.value = shown;
          }} />
      </td>
    );
  };

  // Nhấp đúp / F2: vào sửa trong ô – ô công thức hiện CÔNG THỨC, ô giá trị hiện GIÁ TRỊ (như Excel)
  const enterEdit = (el: HTMLInputElement, f: string | null) => {
    el.classList.remove('rdy');
    if (f) { el.value = f; el.dataset.orig = f; }
    requestAnimationFrame(() => { el.setSelectionRange(el.value.length, el.value.length); widenEdit(el); });
  };
  // Đang sửa mà chữ / công thức dài hơn ô → ô sửa tràn sang phải đè lên ô bên cạnh (như Excel)
  const widenEdit = (el: HTMLInputElement) => {
    const td = el.closest('td') as HTMLElement | null; if (!td) return;
    if (el.classList.contains('rdy')) { unwidenEdit(el); return; }
    const cs = getComputedStyle(el);
    const ctx = (widenEdit as any).ctx || ((widenEdit as any).ctx = document.createElement('canvas').getContext('2d'));
    ctx.font = `${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    const need = Math.ceil(ctx.measureText(el.value).width) + 18;
    const w = td.clientWidth;
    if (need <= w) { unwidenEdit(el); return; }
    if (getComputedStyle(td).position === 'static') { td.style.position = 'relative'; td.dataset.wpos = '1'; }
    el.classList.add('xl-wide');
    el.style.setProperty('width', `${Math.min(need, 900)}px`, 'important'); el.style.setProperty('min-width', `${Math.min(need, 900)}px`, 'important');
  };
  const unwidenEdit = (el: HTMLInputElement) => {
    if (!el.classList.contains('xl-wide')) return;
    el.classList.remove('xl-wide'); el.style.removeProperty('width'); el.style.removeProperty('min-width');
    const td = el.closest('td') as HTMLElement | null;
    if (td?.dataset.wpos) { td.style.position = ''; delete td.dataset.wpos; }
  };
  const rowCacheRef = useRef(new Map<string, { deps: any[]; el: React.ReactElement }>());
  // Đổi cấu trúc cột / chế độ xem → vẽ lại mọi dòng
  const rowGlobal = useMemo(() => ({}), [p1, p2keys.join('|'), frozen, viewVer, ru, structured, JSON.stringify(xcols), activeSheet, d]);
  // Thân bảng chỉ dựng lại khi số liệu đổi – chọn ô / kéo chuột / thanh fx không vẽ lại 400+ dòng (nhanh hơn nhiều)
  const bodyEl = useMemo(() => !d ? null : (
              <tbody>
                {(() => { let idx = -1; const nextCache = new Map<string, { deps: any[]; el: React.ReactElement }>(); const out = viewRows.map(({ r, it }) => {
                  if (it) idx++;
                  const ck = it ? `i${it.id}` : `g${r.sid ?? r.row}`;
                  const deps = it ? [it, edits[it.id], calcs.get(it.id), JSON.stringify(xvals.vals.get(it.id) ?? null), idx, cur === it.id, sel.has(it.id), rowGlobal, r]
                    : [r, JSON.stringify(groupTotals.get(r.row) ?? null), curSid === r.sid, rowGlobal];
                  const hit = rowCacheRef.current.get(ck);
                  if (hit && hit.deps.length === deps.length && hit.deps.every((v, i) => v === deps[i])) { nextCache.set(ck, hit); return hit.el; }
                  const myIdx = idx;
                  const el = ((): React.ReactElement => { const idx = myIdx;
                  if (!it) {
                    return (
                      <tr key={`g${r.sid ?? r.row}`} data-er={r.row} data-sid={r.sid} className={`grp-row ${curSid !== null && curSid === r.sid ? 'cur' : ''}`} onClick={() => { setCur(null); setCurSid(r.sid ?? null); }}
                        onContextMenu={e => { if (r.sid) { e.preventDefault(); setCurSid(r.sid); setMenu({ x: e.clientX, y: e.clientY, sid: r.sid }); } }}>
                        <td className="rn" style={stick(0)}>{r.row}<span className="row-rz" onMouseDown={e => startRowResize(e)} /></td><td style={stick(1)} />
                        {p1.map((c, i) => <td key={c.key} className={c.cls} style={stick(2 + i)}>{c.get(r)}</td>)}
                        {groupTotals.has(r.row) ? (() => { const g = groupTotals.get(r.row)!; return sumCells(g.rev, g.diff, g.ntp, g.xs); })() : <td colSpan={p2w.length} />}
                      </tr>
                    );
                  }
                  const c = calcs.get(it.id)!;
                  const libTitle = it.lib_price
                    ? `Thư viện: ${it.lib_code ? it.lib_code + ' · ' : ''}${it.lib_description}${it.lib_unit ? ' (' + it.lib_unit + ')' : ''}\nVT ${money(it.lib_material) || '—'} · NC ${money(it.lib_labor) || '—'} · Máy ${money(it.lib_machine) || '—'}` +
                      `${it.match_score ? `\nĐộ khớp ${Math.round(it.match_score)}%` : ''}${it.lib_flag === 'uncertain' ? '\nKhớp chưa chắc chắn – kiểm tra tương đương' : ''}`
                    : 'Không tìm thấy trong thư viện';
                  return (
                    <tr key={it.id} data-ri={idx} data-er={it.is_added ? undefined : r.row} className={`st-${STATUS_CLS[c.status]} ${cur === it.id ? 'cur' : ''} ${sel.has(it.id) ? 'sel' : ''} ${edits[it.id] ? 'edited' : ''} ${it.is_added ? 'added' : ''}`} onClick={() => { setCur(it.id); setCurSid(r.sid ?? null); }}
                      onContextMenu={e => { if (r.sid) ctxRef.current(e, r.sid, it); }}>
                      <td className="rn" style={stick(0)} title={it.is_added ? 'Dòng KSCP thêm vào (không có trong file NTP)' : 'Dòng trong file Excel'}>{it.is_added ? '+' : r.row}<span className="row-rz" onMouseDown={e => startRowResize(e)} /></td>
                      <td className="c" style={stick(1)}><input type="checkbox" checked={sel.has(it.id)} onChange={() => setSel(s => { const n = new Set(s); n.has(it.id) ? n.delete(it.id) : n.add(it.id); return n; })} /></td>
                      {p1.map((col, i) => {
                        const af = it.is_added ? ({ stt: 'add_stt', description: 'add_desc', unit: 'add_unit' } as Record<string, 'add_stt' | 'add_desc' | 'add_unit'>)[col.field || ''] : undefined;
                        if (it.is_added) {
                          const v = af ? ((edits[it.id] as any)?.[af] !== undefined ? (edits[it.id] as any)[af] : it[af]) || '' : '';
                          return <td key={col.key} className={`${af ? 'edit note' : ''} ${i === frozen.n ? 'p1s' : ''}`} style={stick(2 + i)}>
                            {af && <input data-r={idx} data-c={af} defaultValue={v} key={`${af}${it.id}-${v}`} onKeyDown={onCellKey} id={af === 'add_desc' ? `add-${it.id}` : undefined}
                              placeholder={af === 'add_desc' ? 'Nhập hạng mục…' : ''} onBlur={e => { const t = e.currentTarget.value.trim(); if (t !== v) editRow(it, { [af]: t || null }); }} />}
                          </td>;
                        }
                        if ((col.field === 'stt' || col.field === 'description') && structured) {  // STT / Mô tả: sửa trực tiếp
                          const tk = col.field === 'stt' ? 'x_t_stt' : 'x_t_desc';
                          const ov = (edits[it.id] as any)?.[tk] !== undefined ? (edits[it.id] as any)[tk] : (it as any)[tk];
                          const orig = String(r.cells[i] ?? '').trim();
                          const shownT = ov ?? orig;
                          return <td key={col.key} className={`${col.cls} edit txt ${i === frozen.n ? 'p1s' : ''} ${ov != null && ov !== orig ? 'ntpchg' : ''}`} style={stick(2 + i)}
                            title={col.field === 'description' ? (ov != null && ov !== orig ? `Mô tả gốc: ${orig}\n` : '') + (libTitle || '') : ov != null && ov !== orig ? `STT gốc: ${orig}` : undefined}>
                            {col.field === 'description'
                              ? <textarea data-r={idx} data-c={tk} defaultValue={shownT} key={`${tk}${it.id}-${shownT}`} rows={1} className="desc-edit" spellCheck={false}
                                  onFocus={e => { setActive(null); setCellKind('v'); }}
                                  onKeyDown={e => { if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); (e.currentTarget as HTMLTextAreaElement).blur(); } else if (e.key === 'Escape') { e.currentTarget.value = shownT; e.currentTarget.blur(); } }}
                                  onBlur={e => { const t = e.currentTarget.value.trim(); if (t === shownT) return; editRow(it, { [tk]: t === orig ? null : t } as any); }} />
                              : <input data-r={idx} data-c={tk} defaultValue={shownT} key={`${tk}${it.id}-${shownT}`} onKeyDown={onCellKey}
                                  onFocus={e => { e.currentTarget.select(); setActive(null); setCellKind('v'); }}
                                  onBlur={e => { const t = e.currentTarget.value.trim(); if (t === shownT) return; editRow(it, { [tk]: t === orig ? null : t } as any); }} />}
                          </td>;
                        }
                        if ((col.field === 'start_date' || col.field === 'end_date') && structured) {  // ngày: sửa trực tiếp dd/mm/yyyy
                          const dk = col.field === 'start_date' ? 'x_d_start' : 'x_d_end';
                          const ov = (edits[it.id] as any)?.[dk] !== undefined ? (edits[it.id] as any)[dk] : (it as any)[dk];
                          const shownD = dateText(ov ?? r.cells[i]);
                          return <td key={col.key} className={`c edit date ${i === frozen.n ? 'p1s' : ''} ${ov ? 'ntpchg' : ''}`} style={stick(2 + i)} title={ov ? `Ngày gốc trong file: ${dateText(r.cells[i])}` : 'Sửa ngày dd/mm/yyyy (thời gian khấu hao tính lại khi bấm Lưu)'}>
                            <input data-r={idx} data-c={dk} defaultValue={shownD} key={`${dk}${it.id}-${shownD}`} onKeyDown={onCellKey}
                              onFocus={e => { e.currentTarget.select(); setActive(null); setCellKind('v'); }}
                              onBlur={e => { const t = e.currentTarget.value.trim(); if (t === shownD) return;
                                if (t && !/^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$/.test(t)) { toast('Nhập ngày dạng dd/mm/yyyy', 'err'); e.currentTarget.value = shownD; return; }
                                editRow(it, { [dk]: t ? dateText(t) : null } as any); }} />
                          </td>;
                        }
                        const nf = NTP_EDIT[col.field || ''];
                        if (nf && structured) {  // ô số NTP: sửa / gõ công thức như Excel
                          const letter = /^c\d+$/.test(col.key) ? colName(Number(col.key.slice(1))) : colName(i + 1);
                          const orig = (() => { try { return JSON.parse(it.rv_extra || '{}')['o_' + nf]; } catch { return undefined; } })();
                          return <React.Fragment key={col.key}>{numCell(it, idx, nf, c.ntp[NTP_VAL[nf]], { qty: nf === 'boq_quantity', pct: nf === 'boq_dep_pct' || nf === 'boq_dep_done', ntp: true, addr: `${letter}${r.row}`, colId: `p1:${col.key}`, fileF: r.f?.[i] ?? null,
                            cls: `${i === frozen.n ? 'p1s' : ''} ${orig !== undefined || (edits[it.id] as any)?.[nf] !== undefined ? 'ntpchg' : ''}`,
                            title: orig !== undefined ? `Số NTP gốc: ${nf === 'boq_dep_pct' ? pctText(orig) : money(orig) || orig}` : undefined })}</React.Fragment>;
                        }
                        if (col.field === 'amount' && structured) {  // Thành tiền NTP: sửa / gõ công thức như Excel (trống = KL × Đơn giá)
                          const letter = /^c\d+$/.test(col.key) ? colName(Number(col.key.slice(1))) : colName(i + 1);
                          const ovSet = typeof ((edits[it.id] as any)?.x_ov_ntp_amt !== undefined ? (edits[it.id] as any).x_ov_ntp_amt : (it as any).x_ov_ntp_amt) === 'number';
                          const chg = !ovSet && Math.abs(c.ntpAmt - (typeof r.cells[i] === 'number' ? r.cells[i] : c.ntpAmt)) > 0.5;
                          return <React.Fragment key={col.key}>{numCell(it, idx, 'x_ov_ntp_amt', c.ntpAmt, { ntp: true, addr: `${letter}${r.row}`, colId: `p1:${col.key}`, fileF: r.f?.[i] ?? null,
                            cls: `${i === frozen.n ? 'p1s' : ''} ${ovSet || chg ? 'ntpchg' : ''}`, title: chg ? 'Tính lại từ số NTP đã sửa (KL × Đơn giá)' : undefined })}</React.Fragment>;
                        }
                        const ntpChanged = col.field === 'amount' && (Math.abs(c.ntpAmt - (typeof r.cells[i] === 'number' ? r.cells[i] : c.ntpAmt)) > 0.5);
                        return <td key={col.key} className={`${col.cls} ${i === frozen.n ? 'p1s' : ''} ${ntpChanged ? 'ntpchg' : ''}`} style={stick(2 + i)} title={col.field === 'description' ? libTitle : ntpChanged ? 'Tính lại từ số NTP đã sửa (KL × Đơn giá)' : undefined}
                          onClick={(col.field === 'price' || col.field === 'amount') && !r.f?.[i] ? () => showFx(it, col.field!) : undefined}>{ntpChanged ? money(c.ntpAmt) : col.get(r, it)}</td>;
                      })}
                      {p2keys.map(k => {
                        let cell: React.ReactNode = null;
                        if (k === 'qty') cell = numCell(it, idx, 'rv_quantity', c.qty, { qty: true });
                        else if (k === 'vt') cell = numCell(it, idx, 'rv_material', c.rv.rv_material);
                        else if (k === 'nc') cell = c.depDone ? <td className="r ro nc-lock" title="Có % ĐÃ KH (vật tư dùng lại) → không tính nhân công, ô ĐG NC bị khóa">—</td> : numCell(it, idx, 'rv_labor', c.rv.rv_labor);
                        else if (k === 'dep') { const ki = d?.kh_info?.[String(it.id)];
                          const kt = ki?.months ? `Thời gian ${ki.months} tháng = ROUNDUP((Kết thúc − Bắt đầu)/30) · ${ki.can ? `vật tư có khấu hao${ki.group ? ' – ' + ki.group : ''}${ki.pct !== null ? ` → bảng ${d?.kh_region === 'gan' ? 'GẦN BIỂN' : 'XA BIỂN'} T=${ki.t}: ${(ki as any).done ? `${pctText((ki as any).base)} − đã KH ${pctText((ki as any).done)} = ` : ''}${pctText(ki.pct)}` : ' (chưa có bảng KH vật tư)'}` : 'mặc định khấu hao 100%'}` : '';
                          cell = numCell(it, idx, 'rv_dep_pct', c.dep, { pct: true, title: (kt ? kt + '\n' : '') + 'Đơn giá = ĐG VT × %KH + ĐG NC (để trống = 100%)' }); }
                        else if (k === 'total') cell = numCell(it, idx, 'rv_total', c.total, c.hasParts
                          ? { title: (c.dep !== null ? 'Tự tính = ĐG VT × %KH + ĐG NC.' : 'Tự tính = ĐG VT + ĐG NC.') + ' Gõ số / công thức vào đây để nhập thẳng đơn giá tổng (bỏ tách VT, NC).' }
                          : { title: 'Nhập đơn giá tổng (dòng không tách VT/NC)' });
                        else if (k === 'amt') cell = numCell(it, idx, 'x_ov_rv_amt', c.amt, { cls: 'strong', title: 'Thành tiền = ĐG tổng × KL · gõ số / công thức để sửa, xóa trắng để về công thức' });
                        else if (k === 'diff') cell = <td onClick={() => showFx(it, 'diff')} className={`r fx ${c.diff < 0 ? 'neg' : c.diff > 0 ? 'pos' : ''}`}>{nfText(c.diff, stylesOf(it)['p2:diff']?.nf) ?? (c.diff ? money(c.diff) : '0')}</td>;
                        else if (k === 'pct') cell = pctCell(k, c.ntpAmt ? (c.diff / c.ntpAmt) * 100 : null, false, it.is_added ? 'Dòng bổ sung – không có giá NTP' : `= ${money(c.diff) || 0} / ${money(c.ntpAmt) || 0}`);
                        else if (k === 'status') cell = <td className="c"><span className={`cx-badge ${STATUS_CLS[c.status]}`}>{c.status}</span></td>;
                        else if (k.startsWith('k3_')) {
                          const k3 = c.k3; const f3 = K3_FIELD[k]; const ownCls = f3 && ((edits[it.id] as any)?.[f3] !== undefined ? (edits[it.id] as any)[f3] : (it as any)[f3]) != null ? 'k3own' : 'k3inh';
                          const t3 = 'P.KSCP soát xét – trống = theo Phần 2 (BCHCT); gõ số / công thức để sửa, xóa trắng để về theo Phần 2';
                          if (k === 'k3_qty') cell = numCell(it, idx, f3, k3.qty, { qty: true, cls: ownCls, title: t3 });
                          else if (k === 'k3_vt') cell = numCell(it, idx, f3, k3.vt, { cls: ownCls, title: t3 });
                          else if (k === 'k3_nc') cell = c.depDone ? <td className="r ro nc-lock" title="Có % ĐÃ KH → không tính nhân công">—</td> : numCell(it, idx, f3, k3.nc, { cls: ownCls, title: t3 });
                          else if (k === 'k3_dep') cell = numCell(it, idx, f3, k3.dep, { pct: true, cls: ownCls, title: t3 });
                          else if (k === 'k3_total') cell = numCell(it, idx, f3, k3.total, { cls: ownCls, title: 'ĐG tổng P.KSCP = ĐG VT × %KH + ĐG NC (theo Phần 3). Gõ số / công thức để nhập thẳng.' });
                          else if (k === 'k3_amt') cell = numCell(it, idx, f3, k3.amt, { cls: 'strong ' + ownCls, title: 'Thành tiền P.KSCP = ĐG tổng × KL · gõ để sửa, xóa trắng để về công thức' });
                          else if (k === 'k3_diff') cell = <td onClick={() => showFx(it, 'k3_diff')} className={`r fx ${k3.diff < 0 ? 'neg' : k3.diff > 0 ? 'pos' : ''}`}>{nfText(k3.diff, stylesOf(it)['p2:k3_diff']?.nf) ?? (k3.diff ? money(k3.diff) : '0')}</td>;
                          else if (k === 'k3_pct') cell = pctCell(k, c.ntpAmt ? (k3.diff / c.ntpAmt) * 100 : null, false, `= ${money(k3.diff) || 0} / ${money(c.ntpAmt) || 0}`);
                          else if (k === 'k3_status') cell = <td className="c"><span className={`cx-badge ${STATUS_CLS[k3.status]}`}>{k3.status}</span></td>;
                          else if (k === 'k3_note') { const nv = (edits[it.id] as any)?.x_k3_note !== undefined ? (edits[it.id] as any).x_k3_note : (it as any).x_k3_note;
                            cell = (<td className={`edit note ${(edits[it.id] as any)?.x_k3_note !== undefined ? 'dirty' : ''}`}>
                              <input data-r={idx} data-c="x_k3_note" defaultValue={nv || ''} key={`k3n${it.id}-${nv || ''}`} onKeyDown={onCellKey}
                                onBlur={e => { const v = e.currentTarget.value.trim(); if (v !== (nv || '')) editRow(it, { x_k3_note: v || null } as any); }} />
                            </td>); }
                        }
                        else if (k === 'note') cell = (
                          <td className={`edit note ${edits[it.id]?.rv_note !== undefined ? 'dirty' : ''}`}>
                            <input data-r={idx} data-c="rv_note" defaultValue={c.rv.rv_note || ''} key={`n${it.id}-${c.rv.rv_note || ''}`} onKeyDown={onCellKey}
                              onBlur={e => { const v = e.currentTarget.value.trim(); if (v !== (c.rv.rv_note || '')) editRow(it, { rv_note: v || null }); }} />
                          </td>);
                        else {
                          const x = xById.get(k); if (!x) return null;
                          const f = k as `x_${string}`;
                          if (x.type === 'formula') {
                            const fv = xvals.vals.get(it.id)?.[k];
                            const bad = typeof fv === 'number' && Number.isNaN(fv);
                            return <td key={k} className={`r fx xf ${bad ? 'neg' : ''}`} title={bad ? xvals.errs.get(k) : `= ${x.formula}`}
                              onClick={() => setFx({ ref: `${it.boq_stt || ''} ${it.boq_description || ''}`.trim(), name: x.label, expr: `= ${x.formula}  →  ${bad ? '#LỖI: ' + (xvals.errs.get(k) || '') : money(fv ?? 0) || '0'}` })}>
                              {bad ? '#LỖI' : fv === null || fv === undefined ? '' : Number.isInteger(fv) || Math.abs(fv) >= 1000 ? money(fv) : num(fv)}</td>;
                          }
                          const v = (edits[it.id] as any)?.[f] !== undefined ? (edits[it.id] as any)[f] : it[f];
                          cell = x.type === 'number' ? numCell(it, idx, f, v ?? null) : (
                            <td className={`edit note ${(edits[it.id] as any)?.[f] !== undefined ? 'dirty' : ''}`}>
                              <input data-r={idx} data-c={f} defaultValue={v || ''} key={`${f}${it.id}-${v || ''}`} onKeyDown={onCellKey}
                                onBlur={e => { const t = e.currentTarget.value.trim(); if (t !== (v || '')) editRow(it, { [f]: t || null }); }} />
                            </td>);
                        }
                        return <React.Fragment key={k}>{cell}</React.Fragment>;
                      })}
                    </tr>
                  );
                  })();
                  nextCache.set(ck, { deps, el }); return el;
                }); rowCacheRef.current = nextCache; return out; })()}
              </tbody>
  ), [d, viewRows, calcs, edits, p1, p2keys.join('|'), frozen, cur, curSid, sel, xvals, groupTotals, structured, activeSheet, viewVer, ru, JSON.stringify(xcols)]);

  if (!d) return <div className="cx-page"><div className="cx-empty">Đang tải gói thầu…</div></div>;
  const allSel = rows.length > 0 && rows.every(r => sel.has(r.id));

  return (
    <div className="cx-page bravo">
      <div className="bv-titlebar">
        <button className="cx-icon-btn" onClick={exit} title="Về danh sách"><Icon name="back" /></button>
        <h1>BẢNG SOÁT XÉT GIÁ ĐIỆN NƯỚC</h1>
        <Badge s={d.status} />
        {dirty && <span className="bv-dirty">● Chưa lưu{Object.keys(edits).length ? ` (${Object.keys(edits).length} dòng)` : ''}</span>}
        <div className="cx-actions">
          <button className="bv-btn orange" onClick={exportOriginal} disabled={!d.session_id} title="File Excel gốc của NTP + các cột KSCP–MEP soát xét bên phải"><Icon name="excel" size={14} /> Xuất Excel</button>
          <button className="bv-btn green" onClick={() => save()} disabled={!dirty || busy === 'save' || !!viewVer}><Icon name="save" size={14} /> {busy === 'save' ? 'Đang lưu…' : 'Lưu'}</button>
          <button className="bv-btn" onClick={exit}>Thoát</button>
        </div>
      </div>

      <div className="bv-form2">
        <div className="r1">
          <label>Ngày tạo</label><input className="ro" readOnly value={dmy(d.created_at)} />
          <label>Dự án</label>
          <div className="bv-loc bv-pa"><input className="ro" readOnly value={d.project_name} title={d.project_name} />
            <span className="bv-heso"><label title="Phương án: PA1 Dự toán (tool áp giá E, F và %KH H từ thư viện) · PA2 BCHCT soát xét · PA3 P.KSCP soát xét">PA</label>
              <select value={pa} disabled={!!viewVer} onChange={async e => {
                const v = e.target.value;
                try { await mepApi(`/api/mep/packages/${id}`, { method: 'PUT', body: JSON.stringify({ pa: v }) });
                  setD(x => (x ? { ...x, pa: v } as any : x)); toast(`Phương án: ${PAS.find(x => x[0] === v)?.[1]}${v === 'PA1' ? ' – bấm Run (Mặc định / Khấu hao) để áp giá E, F và %KH H' : ''}`);
                } catch (er: any) { toast(er.message, 'err'); }
              }}>{PAS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select></span></div>
          <label>Gói thầu</label><input value={form.name} onChange={e => setF('name', e.target.value)} title={d.file_name ? `File báo giá NTP: ${d.file_name}` : ''} />
          <label>Trạng thái</label>
          <select value={form.status} onChange={e => setF('status', e.target.value)}>{(STATUSES.includes(form.status) ? STATUSES : [form.status, ...STATUSES]).map(s => <option key={s}>{s}</option>)}</select>
        </div>
        <div className="r2">
          <label>{pa === 'PA1' ? 'Giá trị dự toán (chưa VAT)' : 'Giá trị NTP báo (chưa VAT)'}</label><input className="ro num" readOnly value={money(totals.ntp)} />
          <label>{pa === 'PA3' ? 'Tổng giá P.KSCP soát xét (chưa VAT)' : pa === 'PA2' ? 'Tổng giá BCHCT soát xét (chưa VAT)' : 'Tổng giá soát xét (chưa VAT)'}</label><input className="ro num strong" readOnly value={pa === 'PA1' ? '—' : money(totals.rev)} />
          <label>Chênh lệch</label><input className={`ro num ${pa !== 'PA1' && totals.diff < 0 ? 'neg' : pa !== 'PA1' && totals.diff > 0 ? 'pos' : ''}`} readOnly value={pa === 'PA1' ? '—' : `${money(totals.diff)}   (${pct(totals.pct)})`} />
          <label title="Vị trí dự án – %KH soát xét tra theo bảng KH VẬT TƯ Xa biển / Gần biển">Vị trí</label>
          <div className="bv-loc"><select value={d.kh_region || 'xa'} disabled={!!viewVer} title="Đổi vị trí: %KH soát xét (ô chưa sửa tay) tính lại theo bảng tương ứng" onChange={async e => {
            try { const r = await mepApi<any>(`/api/mep/projects/${d.project_id}/region`, { method: 'PUT', body: JSON.stringify({ kh_region: e.target.value }) });
              toast(`Vị trí: ${e.target.value === 'gan' ? 'GẦN BIỂN' : 'XA BIỂN'} – tính lại %KH ${r.updated} ô`); await load(); } catch (er: any) { toast(er.message, 'err'); }
          }}><option value="xa">XA BIỂN</option><option value="gan">GẦN BIỂN</option></select>
          <span className="bv-heso"><label title="Hệ số của gói (vd vận hành 2 tháp ×1,8) – nhân vào đơn giá khi Run 'Đơn giá nhân công trực'">Hệ số</label>
            <input className="num" value={form.he_so} onChange={e => setF('he_so', e.target.value)} disabled={!!viewVer} /></span></div>
        </div>
      </div>

      <div className="bv-toolbar">
        <button className="bv-btn teal" onClick={async () => {
          if (!confirm('Tải dữ liệu = về HỒ SƠ GỐC đã tải lên: bỏ hết số / công thức / định dạng đã sửa, dòng đã thêm.\nCác lần Lưu trước vẫn xem lại được ở tab Lịch sử chỉnh sửa. Tiếp tục?')) return;
          setEdits({}); try { if (draftKey) localStorage.removeItem(draftKey); } catch { /* bỏ qua */ }
          await run({ reset: true }, false);
        }} disabled={!d.workbook_id || !!busy} title="Về hồ sơ gốc đã tải lên (bỏ hết chỉnh sửa) – đọc lại file NTP và đối chiếu thư viện"><Icon name="refresh" size={14} /> {busy === 'run' ? 'Đang soát xét…' : 'Tải dữ liệu'}</button>
        <button className="bv-btn teal sq" onClick={() => setTab(tab === 'mapping' ? 'detail' : 'mapping')} disabled={!d.workbook_id} title="Chọn cột trong file NTP"><Icon name="columns" size={14} /></button>
        <button className="bv-btn teal" onClick={() => fileRef.current?.click()} disabled={!!busy}>Nhập dữ liệu từ excel</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xlsm" hidden onChange={e => { const f = e.target.files?.[0]; if (f) importFile(f); e.target.value = ''; }} />
        <span className="bv-sep" />
        <select className="bv-src" value={src} onChange={e => setSrc(e.target.value)} title={SOURCES.find(x => x[0] === src)?.[2]}>
          {SOURCES.map(([k, l, t]) => <option key={k} value={k} title={t}>{l}</option>)}
        </select>
        {(src === 'ncc' || src === 'hd') && (
          <select className="bv-src file" value={srcVer} onChange={e => setSrcVer(e.target.value ? Number(e.target.value) : '')} title="Chọn file đã nạp ở Thư viện giá">
            {!libVers.length && <option value="">— chưa nạp file —</option>}
            {libVers.map(v => <option key={v.id} value={v.id}>{verLabel(v)} · {v.name}</option>)}
          </select>
        )}
        <button className="bv-btn green" onClick={runSource} disabled={!items.length || running || !!viewVer} title="Áp nguồn đã chọn cho ô / vùng đang chọn (chưa chọn gì → cả bảng)">{running ? 'Running…' : '▶ Run'}</button>
        <div className="bv-filter">
          <select value={filter} onChange={e => setFilter(e.target.value)}>
            {FILTERS.map(([k, l]) => <option key={k} value={k}>{l} ({k === '' ? items.length : k.startsWith('lib:') ? libCounts[k.slice(4)] || 0 : k === 'edited' ? libCounts.edited : totals.counts[k] || 0})</option>)}
          </select>
          <div className="cx-search sm"><Icon name="search" size={13} /><input placeholder="Tìm trong bảng…" value={q} onChange={e => setQ(e.target.value)} /></div>
        </div>
      </div>

      <div className="bv-tabs">
        <button className={tab === 'detail' ? 'on' : ''} onClick={() => setTab('detail')}>Chi tiết</button>
        <button className={tab === 'mapping' ? 'on' : ''} onClick={() => setTab('mapping')} disabled={!d.workbook_id}>Cấu hình cột</button>
        <button className={tab === 'history' ? 'on' : ''} onClick={() => setTab('history')}>Lịch sử chỉnh sửa</button>
        <span className="bv-legend">
          {(['Hợp lý', 'Cao', 'Thấp'] as const).map(s => <button key={s} className={`cx-badge ${STATUS_CLS[s]} ${filter === s ? 'on' : ''}`} onClick={() => setFilter(filter === s ? '' : s)}>{s}: {totals.counts[s]}</button>)}
        </span>
      </div>

      {tab === 'mapping' && d.workbook_id && <MappingPanel pkg={d} initialSheet={activeSheet} busy={busy === 'run'} onRun={body => run(body)} onColumnsSaved={() => load(true)} />}
      {tab === 'history' && <HistoryPanel pid={id} onView={async v => { try { const full = await mepApi<SaveVer>(`/api/mep/packages/${id}/versions/${v.id}`); setViewVer(full); setTab('detail'); } catch (e: any) { toast(e.message, 'err'); } }} />}

      {tab === 'detail' && viewVer && (
        <div className="bv-viewver">
          <b>Đang xem lần lưu #{viewVer.no}</b> – {new Date(viewVer.saved_at).toLocaleString('vi-VN', { hour12: false })}{viewVer.saved_by ? ` · ${viewVer.saved_by}` : ''} · chỉ xem, không sửa
          <span className="cx-actions">
            {viewVer.has_file !== 0 && <button className="bv-btn" onClick={() => mepDownload(`/api/mep/packages/${id}/versions/${viewVer.id}/file`).catch(e => toast(e.message, 'err'))}>Tải file Excel bản này</button>}
            <button className="bv-btn green" onClick={() => setViewVer(null)}>Về bản hiện tại</button>
          </span>
        </div>
      )}
      {tab === 'detail' && d.session_id && (() => {
        const as = anchorStyle();
        const tb = (label: React.ReactNode, title: string, on: () => void, active?: boolean, cls = '') =>
          <button type="button" className={`xl-b ${active ? 'on' : ''} ${cls}`} title={title} onMouseDown={e => e.preventDefault()} onClick={on} disabled={!!viewVer}>{label}</button>;
        return (
          <div className="xl-toolbar">
            <div className="xl-grp">
              {tb('↶', 'Hoàn tác (Ctrl+Z)', undo)}{tb('↷', 'Làm lại (Ctrl+Y)', redo)}
            </div>
            <div className="xl-grp">
              {tb('⧉ Copy', 'Copy vùng chọn (Ctrl+C) – kèm công thức', () => copySel())}
              {tb('📋 Dán', 'Dán vào ô đang chọn (Ctrl+V)', toolbarPaste)}
              {tb('⇣ Điền xuống', 'Chép dòng đầu vùng chọn xuống các dòng dưới (Ctrl+D) – hoặc kéo ô vuông góc ô', () => { const n = fillDown(); toast(n ? `Đã điền xuống ${n} ô – bấm Lưu để ghi` : 'Chọn vùng nhiều dòng trước'); })}
            </div>
            <div className="xl-grp">
              {tb(<b>B</b>, 'Đậm (Ctrl+B)', () => { const on = !as.b; applyStyle(s => ({ ...s, b: on })); }, as.b)}
              {tb(<i>I</i>, 'Nghiêng (Ctrl+I)', () => { const on = !as.i; applyStyle(s => ({ ...s, i: on })); }, as.i)}
              {tb(<u>U</u>, 'Gạch chân (Ctrl+U)', () => { const on = !as.u; applyStyle(s => ({ ...s, u: on })); }, as.u)}
              {(['c', 'bg'] as const).map(kind => (
                <span key={kind} className="xl-pal-wrap">
                  <button type="button" className={`xl-b xl-pal-btn ${colorPop === kind ? 'on' : ''}`} title={kind === 'c' ? 'Màu chữ' : 'Tô màu nền ô'} disabled={!!viewVer}
                    onMouseDown={e => e.preventDefault()} onClick={() => setColorPop(colorPop === kind ? null : kind)}>
                    <span className="xl-pal-ic" style={{ borderBottomColor: (kind === 'c' ? as.c : as.bg) || (kind === 'c' ? '#000000' : 'transparent') }}>{kind === 'c' ? 'A' : '▧'}</span> ▾
                  </button>
                  {colorPop === kind && <div className="xl-pal" onMouseDown={e => e.preventDefault()}>
                    <div className="xl-pal-t">{kind === 'c' ? 'Màu chữ' : 'Tô màu ô'}</div>
                    <div className="xl-pal-row">{BASIC_COLORS.map(([cl, nm]) => <button key={cl} type="button" className={`xl-pal-sw ${(kind === 'c' ? as.c : as.bg) === cl ? 'on' : ''}`} style={{ background: cl }} title={nm}
                      onClick={() => { applyStyle(s => ({ ...s, [kind]: cl })); setColorPop(null); }} />)}</div>
                    <button type="button" className="xl-pal-none" onClick={() => { applyStyle(s => { const n = { ...s }; delete n[kind]; return n; }); setColorPop(null); }}>{kind === 'c' ? 'Tự động (đen)' : 'Không tô màu'}</button>
                  </div>}
                </span>))}
            </div>
            <div className="xl-grp">
              {tb('⯇≡', 'Căn trái', () => applyStyle(s => ({ ...s, al: 'l' })), as.al === 'l')}
              {tb('≡', 'Căn giữa', () => applyStyle(s => ({ ...s, al: 'c' })), as.al === 'c')}
              {tb('≡⯈', 'Căn phải', () => applyStyle(s => ({ ...s, al: 'r' })), as.al === 'r')}
            </div>
            <div className="xl-grp">
              <select className="xl-sel" title="Định dạng số" value={as.nf || ''} disabled={!!viewVer} onChange={e => { const v = (e.target.value || undefined) as CellStyle['nf']; applyStyle(s => ({ ...s, nf: v })); }}>
                <option value="">Số: mặc định</option><option value="int">1.234</option><option value="d2">1.234,00</option><option value="pct">% (0,2 → 20%)</option>
              </select>
              {tb('⌫ Xóa định dạng', 'Bỏ đậm, màu, căn lề, định dạng số của vùng chọn', () => applyStyle(() => ({})))}
            </div>
            <span className="xl-hint">{range() ? `Vùng: ${(() => { const g = range()!; return `${g.r2 - g.r1 + 1} dòng × ${g.c2 - g.c1 + 1} cột`; })()}` : 'Bấm / kéo chọn ô để định dạng'}</span>
          </div>
        );
      })()}
      {tab === 'detail' && d.session_id && (
        <div className="bv-fx" title="Bấm vào ô Đơn giá, Thành tiền, Chênh lệch để xem công thức">
          <span className="bv-fx-ref">{active ? active.addr || 'Ô đang chọn' : fx ? fx.name : 'Ô'}</span>
          {cellKind && <span className={`bv-kind k-${cellKind}`} title={cellKind === 'f' ? 'Ô này là CÔNG THỨC – kết quả tự tính lại khi ô nguồn đổi' : cellKind === 'a' ? 'Ô tự tính theo công thức của tool (không nhập)' : 'Ô này là GIÁ TRỊ nhập / số của file'}>
            {cellKind === 'f' ? 'ƒ Công thức' : cellKind === 'a' ? '⚙ Tự tính' : '123 Giá trị'}</span>}
          <b>fx</b>
          {active ? (
            <input className="bv-fx-input" value={fxText} placeholder="Gõ số hoặc công thức, ví dụ =M7*1,1  =SUM(E7:E12) – bấm vào ô trên bảng để chèn địa chỉ"
              onChange={e => { setFxText(e.target.value); const el = e.target; setTimeout(() => updateHint(el), 0); }}
              onBlur={() => setTimeout(() => { setHint(null); const a = document.activeElement as HTMLInputElement | null; if (!isFormulaText(a?.value || '')) clearRefs(); }, 150)}
              onKeyDown={e => {
                if (hintKeyRef.current(e)) return;
                if (e.key === 'Enter' && builtinRef.current && fxText.trim() === builtinRef.current) { e.preventDefault(); setActive(null); return; }
                if (e.key === 'Enter') { e.preventDefault(); if (commitCell(active.it, active.field, fxText, active.pct)) { (document.querySelector(`.cx-grid tr[data-ri] input[data-c="${active.field}"]`) as HTMLElement | null)?.blur(); setActive(null); } }
                else if (e.key === 'Escape') { setActive(null); }
              }} />
          ) : <span className="bv-fx-expr">{fx ? <>{fx.expr}<small>{fx.ref}</small></> : <span className="cx-muted">Bấm vào ô để xem / gõ công thức. Trong ô vàng gõ = rồi bấm vào ô khác để lấy địa chỉ (như Excel)</span>}</span>}
        </div>
      )}
      {tab === 'detail' && (
        <div className="bv-grid-wrap">
          {!d.workbook_id && <div className="cx-empty">Gói chưa có file báo giá NTP. Bấm <b>Nhập dữ liệu từ excel</b> để nạp.</div>}
          {d.workbook_id && !d.session_id && <div className="cx-empty">Chưa soát xét. Mở tab <b>Cấu hình cột</b> để chọn cột rồi bấm Soát xét.</div>}
          {d.session_id && (
            <>
            {sheetNames.length > 1 && (
              <div className="bv-sheet-tabs">
                {sheetNames.map(n => {
                  const its = items.filter((i: any) => i.sheet_name === n);
                  const down = its.filter(i => calcs.get(i.id)!.status === 'Cao').length;
                  return <button key={n} className={n === activeSheet ? 'on' : ''} onClick={() => { setSheet(n); setSel(new Set()); }}>
                    {n} <small>{its.length} dòng{down ? ` · ${down} cao` : ''}</small></button>;
                })}
              </div>
            )}
            {!structured && <div className="bv-hint">Gói này soát xét bằng bản cũ nên Phần 1 đang hiện theo cột chuẩn. Bấm <b>Tải dữ liệu</b> để hiện đúng cột, đúng dòng như file gốc.</div>}
            <style>{`.cx-grid.rv thead tr:not(.grp) > :nth-child(${3 + p1.length}), .cx-grid.rv tbody tr > :nth-child(${3 + p1.length}) { border-left: 4px solid #1d4ed8 !important; }`}</style>
            <table ref={gridRef} tabIndex={-1} onMouseDown={onGridMouseDown} onMouseOver={onGridMouseOver} onPaste={onGridPaste} onKeyDown={onGridKeyDown}
              onInput={e => { const t = e.target as HTMLInputElement; if (t.dataset?.c) { updateHint(t); setFxText(t.value); widenEdit(t); } }}
              onKeyUp={e => { const t = e.target as HTMLInputElement; if (t.dataset?.c && !['ArrowDown', 'ArrowUp', 'Enter', 'Tab', 'Escape'].includes(e.key)) updateHint(t); }}
              onBlur={() => setTimeout(() => { const a = document.activeElement as HTMLInputElement | null; if (!a || !(a.dataset?.c || a.classList?.contains('bv-fx-input'))) setHint(null); if (!isFormulaText(a?.value || '')) clearRefs(); }, 150)}
              onDragStart={e => e.preventDefault()} onDrop={e => e.preventDefault()} onDragOver={e => e.preventDefault()}
              className={`cx-grid rv ${viewVer ? 'viewing' : ''}`} style={{ minWidth: 72 + p1.reduce((a, c) => a + c.width, 0) + p2w.reduce((a, b) => a + b, 0) }}>
              <colgroup>
                <col style={{ width: 40 }} /><col style={{ width: 32 }} />
                {p1.map(c => <col key={c.key} style={{ width: c.width }} />)}
                {p2w.map((w, i) => <col key={'p2' + i} style={{ width: w }} />)}
              </colgroup>
              <thead>
                <tr className="letters">
                  <th style={stick(0, true)} /><th style={stick(1, true)} />
                  {p1.map((c, i) => <th key={c.key} style={stick(2 + i, true)} title="Bấm để chọn cả cột"
                    onMouseDown={e => onLetterDown(2 + i, e)} onMouseEnter={e => onLetterEnter(2 + i, e)}
                    onContextMenu={e => e.preventDefault()}>{c.key.startsWith('c') && /^c\d+$/.test(c.key) ? colName(Number(c.key.slice(1))) : colName(i + 1)}
                    <span className="col-rz" title="Kéo để đổi độ rộng cột" onMouseDown={e => startColResize(e, 2 + i)} onDoubleClick={e => { e.stopPropagation(); setColW(cw => { const n = { ...cw }; delete n['p1:' + c.key]; saveLayout(n); return n; }); }} /></th>)}
                  {p2keys.map((k, i) => <th key={k} title="Bấm để chọn cả cột"
                    onMouseDown={e => onLetterDown(2 + p1.length + i, e)} onMouseEnter={e => onLetterEnter(2 + p1.length + i, e)}
                    onContextMenu={e => e.preventDefault()}>{colName(p2Start + i)}
                    <span className="col-rz" title="Kéo để đổi độ rộng cột" onMouseDown={e => startColResize(e, 2 + p1.length + i)} onDoubleClick={e => { e.stopPropagation(); setColW(cw => { const n = { ...cw }; delete n['p2:' + k]; saveLayout(n); return n; }); }} /></th>)}
                </tr>
                <tr className="grp">
                  <th colSpan={2 + frozen.n} className="stick-l" style={{ position: 'sticky', left: 0, zIndex: 6 }}>{activeSheet} · {sheetItems.length} dòng</th>
                  {p1.length - frozen.n > 0 && <th colSpan={p1.length - frozen.n} className="p1">{(BAND[pa] || BAND.PA2)[0]}</th>}
                  {p2w.length - n3 > 0 && <th colSpan={p2w.length - n3} className="p2">{BAND[pa]?.[1] || BAND.PA2[1]}</th>}
                  {n3 > 0 && <th colSpan={n3} className="p3">{BAND.PA3[2]}</th>}
                </tr>
                <tr>
                  <th style={stick(0, true)} title="Dòng trong file Excel">Dòng</th>
                  <th style={stick(1, true)}><input type="checkbox" checked={allSel} onChange={() => setSel(allSel ? new Set() : new Set(rows.map(r => r.id)))} /></th>
                  {p1.map((c, i) => <th key={c.key} style={stick(2 + i, true)} className={i === frozen.n ? 'p1s' : ''} title={c.label + (c.field ? '' : ' – cột của file gốc, chỉ hiển thị')}><span className="th-clamp">{c.label || '—'}</span></th>)}
                  {p2keys.map(k => {
                    const sd = stdOf(k); const x = xById.get(k);
                    return sd ? <th key={k} className={`${sd.edit ? 'edit-h' : ''} ${k.startsWith('k3_') ? 'k3h' : ''}`} title={sd.title || sd.desc}>{sd.label}{sd.edit ? ' ✎' : ''}</th>
                      : x?.type === 'formula' ? <th key={k} title={`Công thức: = ${x.formula}`}>{x.label} <i className="fx-i">fx</i></th>
                      : <th key={k} className="edit-h" title={`Cột tự tạo (${x?.type === 'number' ? 'số' : 'chữ'}) – sửa trong tab Cấu hình cột`}>{x?.label} ✎</th>;
                  })}
                </tr>
              </thead>
              {bodyEl}
              {!hasGrand && <tfoot>
                <tr>
                  <td colSpan={2 + frozen.n} className="r" style={{ position: 'sticky', left: 0, zIndex: 3 }}>Cộng sheet ({sheetItems.length} dòng)</td>
                  {p1.length - frozen.n > 0 && <td colSpan={p1.length - frozen.n} className="r">NTP: {money(sheetTotals.ntp)}</td>}
                  {sumCells(sheetTotals.rev, sheetTotals.diff, sheetTotals.ntp, sheetTotals.xs)}
                </tr>
              </tfoot>}
            </table>
            </>
          )}
        </div>
      )}
      {hint && (
        <div className="bv-hintbox" style={{ left: hint.x, top: hint.y }} onMouseDown={e => e.preventDefault()}>
          {hint.items.map((f, i) => (
            <button key={f.name} className={i === hint.sel ? 'on' : ''} onClick={() => applyHint(hint, i)}>
              <b>{f.name}</b><span>{f.desc}</span>
            </button>
          ))}
          {hint.sig && <div className="bv-hint-sig"><b>{hint.sig.syntax}</b><span>{hint.sig.desc}</span></div>}
          {hint.items.length > 0 && <small>↑ ↓ chọn · Tab / Enter chèn · Esc đóng</small>}
        </div>
      )}
      {menu && (
        <div className="bv-ctx" style={{ left: menu.x, top: menu.y }} onClick={e => e.stopPropagation()}>
          {clipRef.current?.rowsData?.length ? <button className="strong" onClick={() => { setMenu(null); insertCopiedRows(menu.sid); }}>Chèn {clipRef.current.rowsData.length} dòng đã copy (xuống dưới)</button> : null}
          <button onClick={() => { setMenu(null); addRow('after', menu.sid); }}>Chèn dòng trống phía dưới</button>
          <button onClick={() => { setMenu(null); addRow('before', menu.sid); }}>Chèn dòng trống phía trên</button>
          {menu.it?.is_added ? <button className="danger" onClick={() => { const it = menu.it!; setMenu(null); delRow(it); }}>Xóa dòng này</button> : null}
          <hr />
          <button onClick={() => { setMenu(null); copySel(); }}>Copy (Ctrl+C)</button>
          <button onClick={() => { setMenu(null); toolbarPaste(); }}>Dán (Ctrl+V){clipRef.current?.cells.some(r => r.some(x => x.formula)) ? ' – kèm công thức' : ''}</button>
          <button onClick={() => { setMenu(null); toolbarPaste(true); }}>Dán giá trị (bỏ công thức)</button>
          <button onClick={() => { setMenu(null); const g = range(); if (!g) return; const res = writeCells(g.r1, g.c1, Array.from({ length: g.r2 - g.r1 + 1 }, () => Array(g.c2 - g.c1 + 1).fill(''))); if (res.n) toast(`Đã xóa ${res.n} ô – bấm Lưu để ghi`); }}>Xóa nội dung (Delete)</button>
        </div>
      )}
      <div className="bv-status">
        <span>Dòng: {rows.length}/{items.length}</span>
        <span>Đã chọn: {sel.size}</span>
        <span>Dòng sửa chưa lưu: {Object.keys(edits).length}</span>
        {selStats && <span className="bv-selstats">Vùng chọn: Tổng <b>{money(selStats.sum) || '0'}</b> · TB <b>{num(selStats.avg)}</b> · Đếm <b>{selStats.count}</b> · Min <b>{money(selStats.min) || '0'}</b> · Max <b>{money(selStats.max) || '0'}</b></span>}
        {!selStats && <span className="cx-muted">Ô vàng sửa được · gõ = để tính như Excel (=15.000*1,1) · Enter/↓ xuống dòng · ↑ lên dòng · gõ số kiểu 1.250.000 · để trống = không có</span>}
      </div>
    </div>
  );
}

function HistoryPanel({ pid, onView }: { pid: number; onView: (v: SaveVer) => void }) {
  const [rows, setRows] = useState<Hist[] | null>(null);
  const [vers, setVers] = useState<SaveVer[] | null>(null);
  useEffect(() => {
    mepApi<Hist[]>(`/api/mep/packages/${pid}/history`).then(setRows).catch(e => toast(e.message, 'err'));
    mepApi<SaveVer[]>(`/api/mep/packages/${pid}/versions`).then(setVers).catch(() => setVers([]));
  }, [pid]);
  const time = (s: string) => new Date(s).toLocaleString('vi-VN', { hour12: false });
  const val = (v: string | null) => (v === null || v === '' ? '—' : /^-?\d+(\.\d+)?$/.test(v) ? money(Number(v)) : v);
  return (
    <div className="bv-history">
      <h4 className="bv-h4">Các lần lưu <small>– mỗi lần bấm Lưu tạo một bản; bấm Xem để mở lại bảng lúc đó, Tải file để lấy file Excel bản đó</small></h4>
      <table className="cx-table bv-vers">
        <thead><tr><th>Lần</th><th>Thời gian lưu</th><th>Người lưu</th><th className="r">Số ô sửa</th><th className="r">Tổng NTP</th><th className="r">Tổng soát xét</th><th>Ghi chú</th><th /></tr></thead>
        <tbody>{vers?.map(v => (
          <tr key={v.id} onDoubleClick={() => onView(v)} title="Nhấp đúp để xem lại bản này">
            <td className="strong">#{v.no}</td><td className="nowrap">{time(v.saved_at)}</td><td>{v.saved_by || '—'}</td>
            <td className="r">{v.changes}</td><td className="r">{money(v.ntp_total)}</td><td className="r strong">{money(v.review_total)}</td>
            <td className="cx-muted">{v.note}</td>
            <td className="nowrap">
              <button className="bv-btn" onClick={() => onView(v)}>Xem</button>{' '}
              {v.has_file ? <button className="bv-btn" onClick={() => mepDownload(`/api/mep/packages/${pid}/versions/${v.id}/file`).catch(e => toast(e.message, 'err'))}>Tải file</button> : null}
            </td>
          </tr>
        ))}</tbody>
      </table>
      {vers && !vers.length && <div className="cx-empty sm">Chưa có lần lưu nào – sửa số liệu rồi bấm Lưu.</div>}
    </div>
  );
}

function MappingPanel({ pkg, busy, onRun, initialSheet, onColumnsSaved }: { pkg: Detail; busy: boolean; onRun: (body: any) => void; initialSheet?: string; onColumnsSaved: () => void }) {
  const [sheet, setSheet] = useState(initialSheet || pkg.sheet_name || pkg.sheets?.[0] || '');
  const savedCfg = (pkg.sheet_configs || []).find(c => c.sheet === sheet);
  const included = new Set((pkg.sheet_configs || []).map(c => c.sheet));
  const [headerRow, setHeaderRow] = useState(savedCfg?.header_row || pkg.header_row || 1);
  const [headers, setHeaders] = useState<Record<string, string>>({});
  const [mapping, setMapping] = useState<Record<string, number | null>>({});
  const [threshold, setThreshold] = useState(pkg.session?.threshold_percent || 10);

  const suggest = async (s: string, hr?: number) => {
    try {
      const r = await mepApi<any>(`/api/mep/packages/${pkg.id}/suggest-mapping`, { method: 'POST', body: JSON.stringify({ sheet_name: s, header_row: hr || 1, force_header_row: !!hr }) });
      setHeaders(r.headers || {}); setMapping(r.mapping || {}); setHeaderRow(r.header_row || 1);
    } catch (e: any) { toast(e.message, 'err'); }
  };
  useEffect(() => {
    if (!sheet) return;
    const saved = savedCfg?.mapping || (sheet === pkg.sheet_name && pkg.mapping_json ? JSON.parse(pkg.mapping_json) : null);
    suggest(sheet, saved ? savedCfg?.header_row || pkg.header_row || undefined : undefined).then(() => { if (saved) setMapping(saved); });
  }, [sheet]);
  // P.KSCP soát xét: danh sách cột Phần 2 (cột chuẩn + cột tự tạo), kéo thả để đổi thứ tự
  type Ent = { k: string; col?: CustomCol };
  const initEnts = (): Ent[] => p2Order(pkg.p2_order, pkg.custom_columns || []).map(k => {
    const col = (pkg.custom_columns || []).find(c => 'x_' + c.id === k);
    return col ? { k, col: { ...col } } : { k };
  });
  const [ents, setEnts] = useState<Ent[]>(initEnts);
  const [savingCols, setSavingCols] = useState(false);
  const [dragI, setDragI] = useState<number | null>(null);
  const [overI, setOverI] = useState<number | null>(null);
  const tmpId = useRef(0);
  useEffect(() => { setEnts(initEnts()); }, [JSON.stringify(pkg.custom_columns || []), JSON.stringify(pkg.p2_order || [])]);
  const [ruSel, setRuSel] = useState<number>(pkg.round_unit ?? 100);
  useEffect(() => { setRuSel(pkg.round_unit ?? 100); }, [pkg.round_unit]);
  const colsDirty = JSON.stringify(ents) !== JSON.stringify(initEnts()) || ruSel !== (pkg.round_unit ?? 100);
  const moveEnt = (from: number, to: number) => setEnts(es => {
    if (to < 0 || to >= es.length || from === to) return es;
    const n = [...es]; const [x] = n.splice(from, 1); n.splice(to, 0, x); return n;
  });
  const setCol = (i: number, patch: Partial<CustomCol>) => setEnts(es => es.map((e, j) => j === i && e.col ? { ...e, col: { ...e.col, ...patch } } : e));
  const addCol = () => { tmpId.current += 1; setEnts(es => [...es, { k: 'tmp' + tmpId.current, col: { id: '', label: '', type: 'number' } }]); };
  // Tên cột dùng được trong công thức
  const refNames = [...REF_STD.map(r => r[0]), ...ents.filter(e => e.col && e.col.type !== 'text' && e.col.label.trim()).map(e => e.col!.label.trim())];
  const insertAt = (i: number, text: string) => {
    const el = document.getElementById(`fx-${i}`) as HTMLInputElement | null;
    const cur = ents[i]?.col?.formula || '';
    const pos = el?.selectionStart ?? cur.length;
    const base = cur || '=';
    const p0 = cur ? pos : 1;
    setCol(i, { formula: base.slice(0, p0) + text + base.slice(p0) });
    setTimeout(() => { if (el) { el.focus(); el.setSelectionRange(p0 + text.length, p0 + text.length); } }, 0);
  };
  const saveCols = async () => {
    const customs = ents.filter(e => e.col);
    if (customs.some(e => !e.col!.label.trim())) { toast('Nhập tên cho mọi cột tự tạo', 'err'); return; }
    const labels = customs.map(e => refKey(e.col!.label));
    if (new Set(labels).size !== labels.length || labels.some(l => REF_STD.some(r => refKey(r[0]) === l))) { toast('Tên cột tự tạo bị trùng với cột khác', 'err'); return; }
    for (const e of customs) {
      if (e.col!.type !== 'formula') continue;
      const err = !e.col!.formula?.trim() ? 'chưa nhập công thức' : checkFormula(e.col!.formula, refNames.filter(n => refKey(n) !== refKey(e.col!.label)));
      if (err) { toast(`Cột "${e.col!.label}": ${err}`, 'err'); return; }
    }
    const columns = customs.map(e => e.col!);
    const order = ents.map(e => (e.col ? (e.col.id ? 'x_' + e.col.id : `new:${customs.indexOf(e)}`) : e.k));
    setSavingCols(true);
    try {
      await mepApi(`/api/mep/packages/${pkg.id}/columns`, { method: 'PUT', body: JSON.stringify({ columns, order, round_unit: ruSel }) });
      toast('Đã lưu cột – xem ở tab Chi tiết'); onColumnsSaved();
    } catch (e: any) { toast(e.message, 'err'); } finally { setSavingCols(false); }
  };
  // Cột NTP xếp theo thứ tự cột trong file (trái → phải); cột chưa dùng xuống cuối
  const ntpFields = [...MAP_FIELDS].sort((a, b) => (mapping[a[0]] || 999) - (mapping[b[0]] || 999));
  const ok = mapping.description && (mapping.price || mapping.material_price || mapping.labor_price || mapping.machine_price);

  return (
    <div className="bv-mapping">
      <div className="bv-map-row">
        <div className="cx-field"><span className="cx-label">Sheet báo giá</span>
          <select value={sheet} onChange={e => setSheet(e.target.value)}>{(pkg.sheets || []).map(s => <option key={s} value={s}>{s}{included.has(s) ? ' ✓ đang soát xét' : ''}</option>)}</select></div>
        <div className="cx-field"><span className="cx-label">Dòng tiêu đề (dòng cuối nếu tiêu đề 2 tầng)</span>
          <input type="number" min={1} value={headerRow} onChange={e => setHeaderRow(Number(e.target.value) || 1)} onBlur={() => suggest(sheet, headerRow)} /></div>
        <div className="cx-field"><span className="cx-label">Ngưỡng chênh lệch với thư viện (%)</span>
          <input type="number" min={1} max={100} value={threshold} onChange={e => setThreshold(Number(e.target.value) || 10)} /></div>
        <button className="cx-btn" onClick={() => suggest(sheet)}><Icon name="refresh" size={14} /> Nhận diện lại</button>
      </div>
      <div className="bv-map-split">
        <section className="bv-map-side ntp">
          <h4>Cấu hình báo giá NTP <small>– cột trong file gốc, xếp theo thứ tự trái → phải của Phần 1</small></h4>
          {(() => { const miss = REQUIRED_COLS.filter(([k]) => !mapping[k as MapKey]); return miss.length
            ? <div className="bv-map-miss">⚠ File chưa có đủ cột theo form chuẩn – thiếu: <b>{miss.map(x => x[1]).join(', ')}</b></div>
            : <div className="bv-map-ok">✓ Đủ cột theo form chuẩn</div>; })()}
          {ntpFields.map(([k, label]) => (
            <div key={k} className="bv-map-line">
              <span className="bv-map-col">{mapping[k] ? colName(mapping[k]!) : '–'}</span>
              <span className="bv-map-name">{label}</span>
              <select value={mapping[k] ?? ''} onChange={e => setMapping({ ...mapping, [k]: e.target.value ? Number(e.target.value) : null })} className={mapping[k] ? 'set' : ''}>
                <option value="">— Không dùng —</option>
                {Object.entries(headers).map(([c, t]) => <option key={c} value={c}>Cột {colName(Number(c))}: {t}</option>)}
              </select>
            </div>
          ))}
        </section>
        <section className="bv-map-side kscp">
          <h4>Cấu hình P.KSCP soát xét <small>– cột Phần 2, xếp theo thứ tự trái → phải</small></h4>
          {ents.map((e, i) => {
            const sd = P2_STD.find(x => x.key === e.k);
            return (
              <div key={e.k} className={`bv-map-line ${e.col ? 'custom' : 'fixed'} drag ${dragI === i ? 'dragging' : ''} ${overI === i && dragI !== null && dragI !== i ? (dragI < i ? 'over-down' : 'over-up') : ''}`}
                draggable onDragStart={ev => { setDragI(i); ev.dataTransfer.effectAllowed = 'move'; ev.dataTransfer.setData('text/plain', String(i)); }}
                onDragOver={ev => { ev.preventDefault(); if (overI !== i) setOverI(i); }}
                onDrop={ev => { ev.preventDefault(); if (dragI !== null) moveEnt(dragI, i); setDragI(null); setOverI(null); }}
                onDragEnd={() => { setDragI(null); setOverI(null); }}>
                <span className="bv-map-handle" title="Kéo để đổi thứ tự">⋮⋮</span>
                <span className="bv-map-col">{i + 1}</span>
                {sd ? <>
                  <span className="bv-map-name">{sd.label}{sd.edit ? ' ✎' : ''}</span>
                  <span className="cx-muted bv-map-desc">{sd.desc}</span>
                </> : <>
                  <input value={e.col!.label} placeholder="Nhập tên cột mới…" maxLength={60} autoFocus={!e.col!.id && !e.col!.label}
                    draggable onDragStart={ev => { ev.preventDefault(); ev.stopPropagation(); }}
                    onChange={ev => setCol(i, { label: ev.target.value })} />
                  <select value={e.col!.type} onChange={ev => setCol(i, { type: ev.target.value as CustomCol['type'] })}>
                    <option value="number">Số / tiền</option><option value="text">Chữ</option><option value="formula">Công thức</option>
                  </select>
                </>}
                <span className="bv-map-btns">
                  <button className="cx-btn sm" title="Lên" onClick={() => moveEnt(i, i - 1)} disabled={i === 0}>↑</button>
                  <button className="cx-btn sm" title="Xuống" onClick={() => moveEnt(i, i + 1)} disabled={i === ents.length - 1}>↓</button>
                  {e.col ? <button className="cx-btn sm danger" title="Xóa cột (dữ liệu đã nhập ở cột này sẽ không hiện nữa)" onClick={() => confirm(`Xóa cột "${e.col!.label || 'mới'}"?`) && setEnts(es => es.filter((_, j) => j !== i))}>✕</button>
                    : <span className="bv-map-lock" title="Cột chuẩn – không xóa được">🔒</span>}
                </span>
                {e.col?.type === 'formula' && (() => {
                  const err = checkFormula(e.col.formula || '', refNames.filter(n => refKey(n) !== refKey(e.col!.label)));
                  return (
                    <div className="bv-map-formula" onDragStart={ev => { ev.preventDefault(); ev.stopPropagation(); }} draggable>
                      <span className="fx-i">fx</span>
                      <input value={e.col.formula || ''} placeholder="= [ĐG VT] * [KL] * 1,1   hoặc   = SUM([ĐG VT];[ĐG NC])" id={`fx-${i}`}
                        className={e.col.formula && err ? 'bad' : ''} onChange={ev => setCol(i, { formula: ev.target.value })} />
                      <small className={err && e.col.formula ? 'neg' : 'cx-muted'}>{e.col.formula ? (err ? `Lỗi: ${err}` : 'Công thức hợp lệ') : 'Bấm tên cột / hàm bên dưới để chèn'}</small>
                      <div className="bv-map-chips">
                        {refNames.filter(n => refKey(n) !== refKey(e.col!.label)).map(n => <button key={n} type="button" onClick={() => insertAt(i, `[${n}]`)}>[{n}]</button>)}
                        {['SUM', 'SUBTOTAL', 'ROUND', 'IF', 'MIN', 'MAX', 'AVERAGE', 'ABS'].map(f => <button key={f} type="button" className="fn" onClick={() => insertAt(i, f === 'SUBTOTAL' ? 'SUBTOTAL(9;' : `${f}(`)}>{f}</button>)}
                      </div>
                    </div>
                  );
                })()}
              </div>
            );
          })}
          <div className="bv-map-actions">
            <span className="cx-muted" style={{ marginRight: 'auto' }}>Kéo ⋮⋮ hoặc bấm ↑ ↓ để đổi thứ tự. Cột kiểu Công thức tính như Excel: {FUNCTION_NAMES.slice(0, 6).join(', ')}…</span>
            <button className="cx-btn" onClick={addCol}>+ Tạo cột</button>
            <button className="cx-btn primary" disabled={!colsDirty || savingCols} onClick={saveCols}>{savingCols ? 'Đang lưu…' : 'Lưu cột'}</button>
          </div>
        </section>
      </div>
      <div className="bv-map-foot">
        <span className="cx-muted">Tool soát xét mọi sheet có bảng giá; mỗi sheet một cấu hình cột. Cần cột <b>Hạng mục</b> và <b>Đơn giá tổng</b> (hoặc VT / NC / Khấu hao – tool tự cộng). Số liệu lấy nguyên từ file.</span>
        <span className="cx-actions">
          {included.has(sheet) && included.size > 1 && <button className="cx-btn" disabled={busy} onClick={() => confirm(`Bỏ sheet "${sheet}" khỏi soát xét?`) && onRun({ remove_sheet: sheet, threshold_percent: threshold })}>Bỏ sheet này</button>}
          <button className="cx-btn primary" disabled={!ok || busy} onClick={() => onRun({ sheet_name: sheet, header_row: headerRow, mapping, threshold_percent: threshold })}>
            {busy ? 'Đang soát xét…' : included.has(sheet) ? 'Soát xét lại với cột đã chọn' : 'Thêm sheet này vào soát xét'}
          </button>
        </span>
      </div>
    </div>
  );
}

function colNum(c: string) { return c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
function colName(n: number) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
