// Bộ tính công thức kiểu Excel cho bảng soát xét (chạy hoàn toàn trên trình duyệt, không dùng eval).
// Hỗ trợ: số, tham chiếu cột [Tên cột], + - * / ^ &, so sánh (= <> < > <= >=), % sau số,
// hàm SUM, SUBTOTAL, AVERAGE, MIN, MAX, ROUND, ROUNDUP, ROUNDDOWN, ABS, IF, AND, OR. Dấu phân cách đối số: , hoặc ;
export type Env = (name: string) => number | null;
/** Ô kiểu Excel: cột chữ (A, M, AB…) + số dòng Excel. */
export type CellEnv = (col: string, row: number) => number | null;

type Tok = { t: 'num' | 'ref' | 'id' | 'op' | 'lp' | 'rp' | 'sep' | 'cell' | 'range'; v: string };

function lex(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (/[0-9.]/.test(ch)) {
      let j = i; while (j < src.length && /[0-9.]/.test(src[j])) j++;
      let v = src.slice(i, j);
      if (/^\d{1,3}(\.\d{3})+$/.test(v)) v = v.replace(/\./g, '');  // 15.000 kiểu Việt Nam = 15000
      out.push({ t: 'num', v }); i = j; continue;
    }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      if (j < 0) throw new Error('Thiếu dấu ] đóng tên cột');
      out.push({ t: 'ref', v: src.slice(i + 1, j).trim() }); i = j + 1; continue;
    }
    if (/[A-Za-z_$]/.test(ch)) {
      let j = i; while (j < src.length && /[A-Za-z0-9_.$]/.test(src[j])) j++;
      const w = src.slice(i, j).toUpperCase().replace(/\$/g, '');
      if (/^[A-Z]{1,3}\d+$/.test(w)) {  // ô Excel M7, vùng M7:M10
        const m = /^:\$?([A-Za-z]{1,3})\$?(\d+)/.exec(src.slice(j));
        if (m) { out.push({ t: 'range', v: `${w}:${m[1].toUpperCase()}${m[2]}` }); i = j + m[0].length; continue; }
        out.push({ t: 'cell', v: w }); i = j; continue;
      }
      out.push({ t: 'id', v: w }); i = j; continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') { out.push({ t: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%'.includes(ch)) { out.push({ t: 'op', v: ch }); i++; continue; }
    if (ch === '(') { out.push({ t: 'lp', v: ch }); i++; continue; }
    if (ch === ')') { out.push({ t: 'rp', v: ch }); i++; continue; }
    if (ch === ',' || ch === ';') { out.push({ t: 'sep', v: ch }); i++; continue; }
    throw new Error(`Ký tự không hợp lệ: ${ch}`);
  }
  return out;
}

const n0 = (v: number | null) => (v === null || Number.isNaN(v) ? 0 : v);
const roundTo = (x: number, d: number, mode: 'r' | 'u' | 'd') => {
  const f = Math.pow(10, d); const y = x * f;
  const r = mode === 'r' ? Math.round(y + Math.sign(y) * 1e-9) : mode === 'u' ? Math.sign(y) * Math.ceil(Math.abs(y) - 1e-9) : Math.sign(y) * Math.floor(Math.abs(y) + 1e-9);
  return r / f;
};
type Arg = number | null | (number | null)[];
const flat = (a: Arg[]) => a.flatMap(x => (Array.isArray(x) ? x : [x]));
const nums = (a: Arg[]) => flat(a).filter(x => x !== null && !Number.isNaN(x)) as number[];
const agg = (code: number, v: number[]) => {
  switch (code % 100) {
    case 1: return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null;
    case 2: case 3: return v.length;
    case 4: return v.length ? Math.max(...v) : 0;
    case 5: return v.length ? Math.min(...v) : 0;
    case 6: return v.reduce((s, x) => s * x, 1);
    case 9: return v.reduce((s, x) => s + x, 0);
    default: throw new Error('SUBTOTAL: mã hàm hỗ trợ 1 (TB), 2 (đếm), 4 (max), 5 (min), 6 (tích), 9 (tổng)');
  }
};
const FUNCS: Record<string, (a: Arg[]) => number | null> = {
  SUM: a => nums(a).reduce((s, x) => s + x, 0),
  SUMPRODUCT: a => {
    const arrs = a.map(x => (Array.isArray(x) ? x : [x]));
    const n = arrs[0]?.length || 0;
    if (arrs.some(x => x.length !== n)) throw new Error('SUMPRODUCT: các vùng phải cùng kích thước');
    let s = 0; for (let i = 0; i < n; i++) s += arrs.reduce((p, x) => p * n0(x[i]), 1);
    return s;
  },
  SUBTOTAL: a => agg(n0(flat([a[0]])[0]), nums(a.slice(1))),
  AVERAGE: a => agg(1, nums(a)),
  COUNT: a => nums(a).length,
  MIN: a => agg(5, nums(a)),
  MAX: a => agg(4, nums(a)),
  PRODUCT: a => agg(6, nums(a)),
  ROUND: a => roundTo(n0(flat(a)[0]), n0(flat(a)[1] ?? 0), 'r'),
  ROUNDUP: a => roundTo(n0(flat(a)[0]), n0(flat(a)[1] ?? 0), 'u'),
  ROUNDDOWN: a => roundTo(n0(flat(a)[0]), n0(flat(a)[1] ?? 0), 'd'),
  ABS: a => Math.abs(n0(flat(a)[0])),
  IF: a => { const f = flat(a); return n0(f[0]) ? (f[1] ?? 0) : (f[2] ?? 0); },
  AND: a => (flat(a).every(x => n0(x)) ? 1 : 0),
  OR: a => (flat(a).some(x => n0(x)) ? 1 : 0),
};
/** Gợi ý hàm khi gõ (tên, cú pháp, giải thích). */
export const FUNC_INFO: { name: string; syntax: string; desc: string }[] = [
  { name: 'SUM', syntax: 'SUM(số1; số2 hoặc vùng…)', desc: 'Cộng các số / vùng, ví dụ =SUM(M7:M12)' },
  { name: 'SUMPRODUCT', syntax: 'SUMPRODUCT(vùng1; vùng2; …)', desc: 'Tổng các tích từng cặp, ví dụ =SUMPRODUCT(D7:D12;E7:E12) = Σ KL × ĐG' },
  { name: 'SUBTOTAL', syntax: 'SUBTOTAL(mã; vùng)', desc: 'Tổng phụ: 9 = tổng, 1 = trung bình, 2 = đếm, 4 = lớn nhất, 5 = nhỏ nhất. Ví dụ =SUBTOTAL(9;Q7:Q20)' },
  { name: 'ROUND', syntax: 'ROUND(số; số_chữ_số)', desc: 'Làm tròn; -2 = tròn trăm (…00), -3 = tròn nghìn. Ví dụ =ROUND(M7*1,05;-2)' },
  { name: 'ROUNDUP', syntax: 'ROUNDUP(số; số_chữ_số)', desc: 'Làm tròn lên' },
  { name: 'ROUNDDOWN', syntax: 'ROUNDDOWN(số; số_chữ_số)', desc: 'Làm tròn xuống' },
  { name: 'IF', syntax: 'IF(điều_kiện; nếu_đúng; nếu_sai)', desc: 'Ví dụ =IF(M7>E7;E7;M7)' },
  { name: 'MIN', syntax: 'MIN(số1; số2 hoặc vùng…)', desc: 'Số nhỏ nhất, ví dụ =MIN(E7;[Giá TV])' },
  { name: 'MAX', syntax: 'MAX(số1; số2 hoặc vùng…)', desc: 'Số lớn nhất' },
  { name: 'AVERAGE', syntax: 'AVERAGE(số1; số2 hoặc vùng…)', desc: 'Trung bình cộng' },
  { name: 'COUNT', syntax: 'COUNT(vùng)', desc: 'Đếm ô có số' },
  { name: 'PRODUCT', syntax: 'PRODUCT(số1; số2…)', desc: 'Nhân các số' },
  { name: 'ABS', syntax: 'ABS(số)', desc: 'Giá trị tuyệt đối' },
  { name: 'AND', syntax: 'AND(đk1; đk2…)', desc: 'Đúng khi mọi điều kiện đúng' },
  { name: 'OR', syntax: 'OR(đk1; đk2…)', desc: 'Đúng khi có một điều kiện đúng' },
];
export const FUNCTION_NAMES = Object.keys(FUNCS);

/** Dấu phẩy thập phân kiểu Việt Nam: có ";" → "," là thập phân; không có hàm "(" → "," là thập phân; còn lại "," tách đối số. */
export function normDecimal(src: string) {
  const s = src.replace(/\[[^\]]*\]/g, m => m.replace(/,/g, '\u0001'));
  const dec = s.includes(';') || !s.includes('(');
  return (dec ? s.replace(/(\d),(\d)/g, '$1.$2') : s).replace(/\u0001/g, ',');
}

const colNum = (c: string) => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colStr = (n: number) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const splitCell = (a: string) => { const m = /^([A-Z]+)(\d+)$/.exec(a)!; return { c: colNum(m[1]), r: Number(m[2]) }; };

/** Làm tròn đơn giá theo bước (100 = …00). */
export const roundUnit = (v: number, unit: number) => (unit > 0 ? Math.round(v / unit + 1e-9 * Math.sign(v)) * unit : v);

export function evalFormula(src: string, env: Env, cellEnv?: CellEnv): number | null {
  const cellVal = (a: string) => {
    if (!cellEnv) throw new Error(`Không dùng được địa chỉ ô ${a} ở đây`);
    const { c, r } = splitCell(a); return cellEnv(colStr(c), r);
  };
  const rangeVals = (v: string) => {
    const [a, b] = v.split(':').map(splitCell);
    const out: (number | null)[] = [];
    for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++)
      for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) { if (!cellEnv) throw new Error('Không dùng được vùng ô ở đây'); out.push(cellEnv(colStr(c), r)); }
    return out;
  };
  const toks = lex(normDecimal(src.replace(/^\s*=/, '')));
  let p = 0;
  const peek = () => toks[p];
  const take = () => toks[p++];
  const cmp = (): number | null => {
    let l = add();
    while (peek()?.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) {
      const op = take().v; const r = add(); const a = n0(l), b = n0(r);
      l = (op === '=' ? a === b : op === '<>' ? a !== b : op === '<' ? a < b : op === '>' ? a > b : op === '<=' ? a <= b : a >= b) ? 1 : 0;
    }
    return l;
  };
  const add = (): number | null => {
    let l = mul();
    while (peek()?.t === 'op' && (peek().v === '+' || peek().v === '-' || peek().v === '&')) {
      const op = take().v; const r = mul();
      l = op === '-' ? n0(l) - n0(r) : n0(l) + n0(r);
    }
    return l;
  };
  const mul = (): number | null => {
    let l = pow();
    while (peek()?.t === 'op' && (peek().v === '*' || peek().v === '/')) {
      const op = take().v; const r = pow();
      if (op === '/') { if (!n0(r)) throw new Error('Chia cho 0'); l = n0(l) / n0(r); } else l = n0(l) * n0(r);
    }
    return l;
  };
  const pow = (): number | null => {
    const l = unary();
    if (peek()?.t === 'op' && peek().v === '^') { take(); return Math.pow(n0(l), n0(pow())); }
    return l;
  };
  const unary = (): number | null => {
    if (peek()?.t === 'op' && (peek().v === '-' || peek().v === '+')) { const op = take().v; const v = unary(); return op === '-' ? -n0(v) : v; }
    let v = atom();
    while (peek()?.t === 'op' && peek().v === '%') { take(); v = n0(v) / 100; }
    return v;
  };
  const atom = (): number | null => {
    const tk = take();
    if (!tk) throw new Error('Công thức chưa hoàn chỉnh');
    if (tk.t === 'num') { const v = Number(tk.v); if (Number.isNaN(v)) throw new Error(`Số không hợp lệ: ${tk.v}`); return v; }
    if (tk.t === 'ref') return env(tk.v);
    if (tk.t === 'cell') return cellVal(tk.v);
    if (tk.t === 'range') { const v = rangeVals(tk.v); if (v.length === 1) return v[0]; throw new Error(`Vùng ${tk.v} phải nằm trong hàm, ví dụ SUM(${tk.v})`); }
    if (tk.t === 'lp') { const v = cmp(); if (take()?.t !== 'rp') throw new Error('Thiếu dấu )'); return v; }
    if (tk.t === 'id') {
      const fn = FUNCS[tk.v];
      if (!fn) throw new Error(`Không có hàm ${tk.v}`);
      if (take()?.t !== 'lp') throw new Error(`Thiếu ( sau ${tk.v}`);
      const args: Arg[] = [];
      const arg = () => {
        const nx = toks[p + 1];
        if (peek()?.t === 'range' && (!nx || nx.t === 'sep' || nx.t === 'rp')) args.push(rangeVals(take().v));
        else args.push(cmp());
      };
      if (peek()?.t !== 'rp') {
        arg();
        while (peek()?.t === 'sep') { take(); arg(); }
      }
      if (take()?.t !== 'rp') throw new Error(`Thiếu ) của ${tk.v}`);
      return fn(args);
    }
    throw new Error(`Không hiểu "${tk.v}"`);
  };
  if (!toks.length) return null;
  const v = cmp();
  if (p < toks.length) throw new Error(`Thừa "${toks[p].v}"`);
  return v === null || !Number.isFinite(v) ? v : Math.round(v * 1e6) / 1e6;
}

/** Tên cột chuẩn hóa để so khớp: bỏ dấu, chữ thường, bỏ khoảng trắng. */
export const refKey = (s: string) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/gi, 'd').toLowerCase().replace(/\s+/g, '');

/** Kiểm tra cú pháp + tên cột (trả lỗi hoặc null). */
export function checkFormula(src: string, names: string[]): string | null {
  const known = new Set(names.map(refKey));
  try {
    evalFormula(src, n => { if (!known.has(refKey(n))) throw new Error(`Không có cột [${n}]`); return 1; });
    return null;
  } catch (e: any) { return e.message === 'Chia cho 0' ? null : e.message; }
}
