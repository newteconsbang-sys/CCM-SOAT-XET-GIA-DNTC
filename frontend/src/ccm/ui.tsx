import React, { useEffect, useRef, useState } from 'react';

export type Project = { id: number; name: string; location: string; packages: { id: number; code: string; name: string; status: string }[] };
export type Version = { id: number; name: string; effective_date: string; scope: string; note: string; status: string; item_count: number; error_count: number; created_by: string; created_at: string; kind?: string };
// Phân loại thư viện giá khi nạp file
export const LIB_KINDS: [string, string][] = [['dntc', 'Đơn giá ĐNTC chuẩn'], ['ncc', 'Báo giá NCC'], ['hd', 'Hợp đồng tham khảo'], ['kh', 'Khấu hao'], ['truc', 'Đơn giá nhân công trực']];
export const libKindLabel = (k?: string) => LIB_KINDS.find(x => x[0] === (k || 'dntc'))?.[1] || 'Đơn giá ĐNTC chuẩn';
export type Pkg = {
  id: number; project_id: number; code: string; name: string; work_type: string; subcontractor: string;
  version_id: number | null; version_name: string; deadline: string; status: string; workbook_id: string | null;
  session_id: string | null; sheet_name: string | null; vat_percent: number | null; note: string | null;
  project_name: string; project_location: string; file_name?: string; created_at: string; updated_at: string;
  item_count: number; ntp_total: number | null; review_total: number | null; diff_percent: number | null;
};

export const WORK_TYPES = ['Điện', 'Nước', 'PCCC', 'Điện nhẹ', 'HVAC'];
export const STATUSES = ['Đang soát xét', 'Đã phản hồi BCHCT', 'Xong'];
// Phương án của gói
export const PAS: [string, string][] = [['PA1', 'PA1: Dự toán ĐNTC'], ['PA2', 'PA2: NTP / BCHCT soát xét'], ['PA3', 'PA3: NTP / BCHCT / P.KSCP']];

const vn0 = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 0 });
const vn2 = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 });
export const money = (v: number | null | undefined) => (v === null || v === undefined || isNaN(v) ? '' : vn0.format(v));
export const num = (v: number | null | undefined) => (v === null || v === undefined || isNaN(v) ? '' : vn2.format(v));
export const million = (v: number | null | undefined) => {
  if (v === null || v === undefined) return '—';
  const m = v / 1e6;
  return new Intl.NumberFormat('vi-VN', { maximumFractionDigits: Math.abs(m) < 100 ? 1 : 0 }).format(m);
};
export const pct = (v: number | null | undefined) =>
  v === null || v === undefined ? '—' : `${v > 0 ? '+' : ''}${new Intl.NumberFormat('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(v)}%`;
export const dmy = (iso?: string | null) => {
  if (!iso) return '';
  const [y, m, d] = iso.slice(0, 10).split('-');
  return d ? `${d}/${m}/${y}` : iso;
};
/** Đọc số kiểu Việt Nam: "1.250.000" hoặc "1250000,5". */
export const parseVn = (s: string): number | null => {
  const t = s.trim().replace(/\s/g, '');
  if (!t) return null;
  const normalized = /,\d{1,2}$/.test(t) ? t.replace(/\./g, '').replace(',', '.') : t.replace(/[.,](?=\d{3}(\D|$))/g, '').replace(',', '.');
  const n = Number(normalized);
  return isNaN(n) ? null : n;
};

const STATUS_CLASS: Record<string, string> = {
  'Xong': 'ok', 'Đã phản hồi BCHCT': 'info', 'Đã duyệt': 'ok', 'Chờ duyệt': 'info', 'Cần giảm': 'bad', 'Đang soát xét': 'warn', 'Nháp': 'muted',
  active: 'ok', archived: 'muted', draft: 'warn', checked: 'info',
};
export const VERSION_LABEL: Record<string, string> = { active: 'Đang áp dụng', archived: 'Lưu trữ', draft: 'Nháp', checked: 'Đã kiểm tra' };
export function Badge({ s, label }: { s: string; label?: string }) {
  return <span className={`cx-badge ${STATUS_CLASS[s] || 'muted'}`}>{label || s}</span>;
}

export function Icon({ name, size = 16 }: { name: string; size?: number }) {
  const p: Record<string, React.ReactNode> = {
    review: <><rect x="4" y="4" width="16" height="16" rx="3" /><path d="m8.5 12 2.5 2.5 5-5" /></>,
    library: <><path d="M6 3h10l3 3v15H6z" /><path d="M9 9h7M9 13h7M9 17h4" /></>,
    grid: <><rect x="3" y="4" width="18" height="16" rx="2" /><path d="M3 10h18M3 15h18M9 4v16M15 4v16" /></>,
    history: <><path d="M3 12a9 9 0 1 0 3-6.7L3 8" /><path d="M3 3v5h5M12 7v5l3 2" /></>,
    settings: <><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.6 1.6 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.6 1.6 0 0 0-1.8-.3 1.6 1.6 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.6 1.6 0 0 0-1-1.5 1.6 1.6 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.6 1.6 0 0 0 .3-1.8 1.6 1.6 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.6 1.6 0 0 0 1.5-1 1.6 1.6 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.6 1.6 0 0 0 1.8.3H9a1.6 1.6 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.6 1.6 0 0 0 1 1.5 1.6 1.6 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.6 1.6 0 0 0-.3 1.8V9a1.6 1.6 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.6 1.6 0 0 0-1.5 1z" /></>,
    download: <><path d="M12 4v11M7 10l5 5 5-5M5 20h14" /></>,
    upload: <><path d="M12 20V9M7 14l5-5 5 5M5 4h14" /></>,
    plus: <><path d="M12 5v14M5 12h14" /></>,
    search: <><circle cx="11" cy="11" r="7" /><path d="m20 20-3.5-3.5" /></>,
    refresh: <><path d="M20 11a8 8 0 1 0-2.3 5.7" /><path d="M20 4v7h-7" /></>,
    swap: <><path d="M7 4 3 8l4 4M3 8h14M17 20l4-4-4-4M21 16H7" /></>,
    close: <><path d="M6 6l12 12M18 6 6 18" /></>,
    send: <><path d="M4 12 20 4l-6 16-3-7z" /></>,
    save: <><path d="M5 4h11l3 3v13H5z" /><path d="M8 4v5h7V4M8 20v-6h8v6" /></>,
    back: <><path d="M15 6l-6 6 6 6" /></>,
    columns: <><path d="M12 4v16M8 8l-4 4 4 4M16 8l4 4-4 4" /></>,
    excel: <><rect x="4" y="3" width="16" height="18" rx="2" /><path d="m9 8 6 8M15 8l-6 8" /></>,
    print: <><path d="M7 9V3h10v6M7 17H4v-7h16v7h-3" /><rect x="7" y="14" width="10" height="7" /></>,
    trash: <><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3" /></>,
    check: <path d="m5 12 5 5 9-10" />,
  };
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      {p[name]}
    </svg>
  );
}

export function Modal({ title, onClose, children, footer, width = 760 }: { title: string; onClose: () => void; children: React.ReactNode; footer?: React.ReactNode; width?: number }) {
  useEffect(() => {
    const k = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', k);
    return () => window.removeEventListener('keydown', k);
  }, [onClose]);
  return (
    <div className="cx-overlay" onMouseDown={e => e.target === e.currentTarget && onClose()}>
      <div className="cx-modal" style={{ width }} role="dialog" aria-label={title}>
        <div className="cx-modal-head">
          <h3>{title}</h3>
          <button className="cx-icon-btn" onClick={onClose} aria-label="Đóng"><Icon name="close" size={18} /></button>
        </div>
        <div className="cx-modal-body">{children}</div>
        {footer && <div className="cx-modal-foot">{footer}</div>}
      </div>
    </div>
  );
}

export function Field({ label, req, children, span }: { label: string; req?: boolean; children: React.ReactNode; span?: number }) {
  return (
    <div className="cx-field" style={span ? { gridColumn: `span ${span}` } : undefined}>
      <span className="cx-label">{label}{req && <b> *</b>}</span>
      {children}
    </div>
  );
}

export function DropZone({ file, onFile, accept = '.xlsx,.xlsm', hint, tall, compact, extra }: { file?: File | null; onFile: (f: File) => void; accept?: string; hint: string; tall?: boolean; compact?: boolean; extra?: React.ReactNode }) {
  const ref = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  if (compact) {
    return (
      <div
        className={`cx-drop compact ${over ? 'over' : ''}`} title={`Kéo thả file vào đây hoặc bấm để chọn · ${hint}`}
        onDragOver={e => { e.preventDefault(); setOver(true); }}
        onDragLeave={() => setOver(false)}
        onDrop={e => { e.preventDefault(); setOver(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
        onClick={() => ref.current?.click()}
      >
        <input ref={ref} type="file" accept={accept} hidden onChange={e => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ''; }} />
        <Icon name="upload" size={14} /> Chọn file
      </div>
    );
  }
  return (
    <div
      className={`cx-drop ${over ? 'over' : ''} ${tall ? 'tall' : ''}`}
      onDragOver={e => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={e => { e.preventDefault(); setOver(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }}
      onClick={() => ref.current?.click()}
    >
      <input ref={ref} type="file" accept={accept} hidden onChange={e => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ''; }} />
      {tall && <div className="cx-drop-icon"><Icon name="upload" size={26} /></div>}
      <div className="cx-drop-main">{!tall && <Icon name="upload" size={14} />} Kéo thả file{tall ? ' Excel' : ''} vào đây hoặc <a>chọn file</a></div>
      <div className="cx-drop-hint">{hint}</div>
      {file && <div className="cx-file-chip" onClick={e => e.stopPropagation()}><Icon name="excel" size={14} /> {file.name}{extra}</div>}
    </div>
  );
}

let toastFn: ((msg: string, kind?: 'ok' | 'err') => void) | null = null;
export const toast = (msg: string, kind: 'ok' | 'err' = 'ok') => toastFn?.(msg, kind);
export function Toaster() {
  const [list, setList] = useState<{ id: number; msg: string; kind: string }[]>([]);
  useEffect(() => {
    toastFn = (msg, kind = 'ok') => {
      const id = Date.now() + Math.random();
      setList(l => [...l, { id, msg, kind }]);
      setTimeout(() => setList(l => l.filter(t => t.id !== id)), kind === 'err' ? 7000 : 3500);
    };
    return () => { toastFn = null; };
  }, []);
  return <div className="cx-toasts">{list.map(t => <div key={t.id} className={`cx-toast ${t.kind}`}>{t.msg}</div>)}</div>;
}

/** Menu ··· gọn cho từng dòng. */
export function RowMenu({ items }: { items: { label: string; onClick: () => void; danger?: boolean; hidden?: boolean }[] }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const h = (e: MouseEvent) => !ref.current?.contains(e.target as Node) && setOpen(false);
    document.addEventListener('mousedown', h);
    return () => document.removeEventListener('mousedown', h);
  }, [open]);
  return (
    <div className="cx-menu-wrap" ref={ref}>
      <button className="cx-btn sm" onClick={() => setOpen(o => !o)} aria-label="Thao tác khác">···</button>
      {open && (
        <div className="cx-menu">
          {items.filter(i => !i.hidden).map(i => (
            <button key={i.label} className={i.danger ? 'danger' : ''} onClick={() => { setOpen(false); i.onClick(); }}>{i.label}</button>
          ))}
        </div>
      )}
    </div>
  );
}

/** Ô sửa trực tiếp: bấm để sửa, Enter/Tab lưu, Esc hủy. onSave trả Promise; lỗi thì giữ ô đang sửa. */
export function EditCell({ value, display, onSave, numeric, options, className, placeholder }: {
  value: string | number | null | undefined; display?: React.ReactNode; onSave: (v: string) => Promise<void> | void;
  numeric?: boolean; options?: string[]; className?: string; placeholder?: string;
}) {
  const [editing, setEditing] = useState(false);
  const [busy, setBusy] = useState(false);
  const ref = useRef<HTMLInputElement & HTMLSelectElement>(null);
  const initial = value === null || value === undefined ? '' : numeric && typeof value === 'number' ? num(value) : String(value);
  useEffect(() => { if (editing) { ref.current?.focus(); ref.current?.select?.(); } }, [editing]);
  const commit = async (v: string) => {
    if (v === initial) { setEditing(false); return; }
    setBusy(true);
    try { await onSave(v); setEditing(false); } catch (e: any) { toast(e.message || String(e), 'err'); ref.current?.focus(); } finally { setBusy(false); }
  };
  if (!editing) {
    return (
      <div className={`cx-edit-cell ${className || ''}`} tabIndex={0} title="Bấm để sửa"
        onClick={() => setEditing(true)} onKeyDown={e => { if (e.key === 'Enter' || e.key === 'F2') setEditing(true); }}>
        {display ?? (initial || <span className="cx-muted">{placeholder || '—'}</span>)}
      </div>
    );
  }
  if (options) {
    return (
      <select ref={ref} className="cx-edit-input" defaultValue={initial} disabled={busy}
        onChange={e => commit(e.target.value)} onBlur={() => setEditing(false)} onKeyDown={e => e.key === 'Escape' && setEditing(false)}>
        {options.map(o => <option key={o}>{o}</option>)}
      </select>
    );
  }
  return (
    <input ref={ref} className={`cx-edit-input ${numeric ? 'r' : ''}`} defaultValue={initial} disabled={busy} inputMode={numeric ? 'decimal' : undefined}
      onBlur={e => commit(e.currentTarget.value.trim())}
      onKeyDown={e => {
        if (e.key === 'Enter') { e.preventDefault(); commit(e.currentTarget.value.trim()); }
        else if (e.key === 'Escape') setEditing(false);
      }} />
  );
}
