import React, { useEffect, useMemo, useRef, useState } from 'react';
import { mepApi, mepDownload } from '../mepApi';
import { Badge, DropZone, EditCell, Field, Icon, LIB_KINDS, Modal, VERSION_LABEL, Version, dmy, libKindLabel, money, num, toast } from './ui';

type Problem = { sheet: string; row: number; code: string; description: string; problems: string[] };
type PriceChange = { code: string; description: string; old?: number; new?: number };
type ImportResult = {
  skipped_sheets?: { sheet: string; reason: string }[];
  version: { id: number; name: string; count: number }; kind?: string; activated?: boolean; filename: string; row_count: number; valid_count: number;
  by_type: Record<string, number>; errors: Problem[]; warnings: Problem[]; changed: PriceChange[]; new: PriceChange[]; removed: PriceChange[];
};
type LibItem = {
  id: number; item_type: string; code: string; description: string; specification: string; brand: string; unit: string; category: string | null; system: string;
  material_price: number | null; labor_price: number | null; machine_price: number | null; waste_percent: number | null; total_price: number;
  compare_percent: number | null; is_new: boolean;
};
type Sheet = { id: number; name: string; header_row: number; first_col: number; last_col: number; headers: string[]; mapping: Record<string, number>; count: number };
type SRow = { id: number; row: number | null; kind: 'item' | 'group' | 'error' | 'deleted' | 'raw' | 'title'; cells: any[]; item?: LibItem };
type Items = { items: LibItem[]; types: { name: string; count: number }[]; categories: { name: string; count: number }[]; compare_name: string | null; compare_to: number | null;
  sheets: Sheet[]; sheet: { sheet_id: number; rows: SRow[] } | null };
type CmpRow = { item_type: string; code: string; description: string; unit: string; material_price: number | null; labor_price: number | null; machine_price: number | null; total_price: number; old?: Record<string, any>; field_labels?: string[]; old_total?: number; diff_percent?: number | null };
type Compare = { base: { id: number; name: string }; target: { id: number; name: string }; added: CmpRow[]; removed: CmpRow[]; changed: CmpRow[]; summary: Record<string, number> };
type Hist = { id: number; action: string; field_label: string; old_value: string | null; new_value: string | null; item_label: string | null; changed_by: string; changed_at: string };

const TYPES = ['Vật tư', 'Nhân công', 'Ca máy'];
const pctTxt = (p: number) => `${p > 0 ? '▲' : '▼'} ${Math.abs(p).toLocaleString('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 2 })}%`;

export function LibraryVersions() {
  const [versions, setVersions] = useState<Version[]>([]);
  const [viewId, setViewId] = useState<number | null>(null);
  const [data, setData] = useState<Items | null>(null);
  const [type, setType] = useState('');
  const [sheetId, setSheetId] = useState<number | null>(null);
  const [cat, setCat] = useState('');
  const [q, setQ] = useState('');
  const [reload, setReload] = useState(0);
  const [file, setFile] = useState<File | null>(null);
  const [kind, setKind] = useState('dntc');  // phân loại file sắp nạp
  const [tab, setTabState] = useState<string>(() => { try { return localStorage.getItem('cx-lib-tab') || 'dntc'; } catch { return 'dntc'; } });  // tab phân loại hoặc 'files'
  const setTab = (t: string) => { setTabState(t); try { localStorage.setItem('cx-lib-tab', t); } catch { /* bỏ qua */ } };
  const kindOf = (v?: Version) => v?.kind || 'dntc';
  // Bản hiển thị mặc định của 1 phân loại: bản đang áp dụng (nếu có), không thì bản mới nhất
  const pickOf = (k: string, vs = versions) => { const xs = vs.filter(v => kindOf(v) === k); return (xs.find(v => v.status === 'active') || xs[0])?.id ?? null; };
  useEffect(() => { if (tab !== 'files') setKind(tab); }, [tab]);
  useEffect(() => {
    if (tab === 'files') return;
    const cur = versions.find(v => v.id === viewId);
    if (!cur || kindOf(cur) !== tab) { setViewId(pickOf(tab)); setSheetId(null); }
  }, [tab, versions]);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<ImportResult | null>(null);
  const [popup, setPopup] = useState<'' | 'errors' | 'warnings' | 'changes'>('');
  const [showCompare, setShowCompare] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [adding, setAdding] = useState(false);

  const loadVersions = async (select?: number) => {
    const v = await mepApi<Version[]>('/api/mep/library/versions');
    setVersions(v);
    setViewId(cur => select ?? (cur && v.some(x => x.id === cur) ? cur : null) ?? v.find(x => x.status === 'active')?.id ?? v[0]?.id ?? null);
  };
  useEffect(() => { loadVersions().catch(e => toast(e.message, 'err')); }, []);
  useEffect(() => {
    if (!viewId) { setData(null); return; }
    const qs = new URLSearchParams({ type, category: cat, q });
    if (sheetId) qs.set('sheet_id', String(sheetId));
    const t = setTimeout(() => mepApi<Items>(`/api/mep/library/versions/${viewId}/items?${qs}`).then(setData).catch(e => toast(e.message, 'err')), q ? 250 : 0);
    return () => clearTimeout(t);
  }, [viewId, type, cat, q, reload, sheetId]);

  const upload = async (f: File) => {
    setFile(f); setResult(null); setBusy(true);
    const fd = new FormData(); fd.append('file', f); fd.append('kind', kind);
    try {
      const r = await mepApi<ImportResult>('/api/mep/library/versions/import', { method: 'POST', body: fd });
      setResult(r);
      toast(`Đã lưu ${r.version.count ? r.version.count.toLocaleString('vi-VN') + ' mục' : 'bảng'} thành ${r.version.name} · ${libKindLabel(r.kind)}${r.activated ? ' và áp dụng' : ''}`);
      setType(''); setCat(''); setQ(''); setSheetId(null);
      setTab(r.kind || kind);
      await loadVersions(r.version.id);
    } catch (e: any) { toast(e.message, 'err'); } finally { setBusy(false); }
  };
  const activate = async (id: number) => {
    const v = versions.find(x => x.id === id);
    if (!confirm(`Đặt ${v?.name} làm version hiện hành? Các gói thầu mới sẽ dùng version này.`)) return;
    try { await mepApi(`/api/mep/library/versions/${id}/activate`, { method: 'POST' }); toast(`${v?.name} đang áp dụng`); loadVersions(id); }
    catch (e: any) { toast(e.message, 'err'); }
  };
  const remove = async (id: number) => {
    const v = versions.find(x => x.id === id);
    const others = versions.filter(x => x.id !== id);
    const msg = `Xóa thư viện giá ${v?.name} (${v?.item_count?.toLocaleString('vi-VN') || 0} mục) và toàn bộ giá trong đó?\n`
      + (v?.status === 'active' ? (others.length ? `Version ${others[0].name} sẽ được đặt làm hiện hành.\n` : 'Sau khi xóa sẽ không còn thư viện giá nào.\n') : '')
      + 'Gói thầu đang dùng version này chuyển sang version còn lại; số liệu đã soát xét của gói giữ nguyên.\nKhông hoàn tác được.';
    if (!confirm(msg)) return;
    try {
      const r = await mepApi<any>(`/api/mep/library/versions/${id}`, { method: 'DELETE' });
      toast(`Đã xóa ${v?.name}${r?.moved_packages ? ` · ${r.moved_packages} gói thầu chuyển sang version khác` : ''}`);
      if (viewId === id) setViewId(null); loadVersions(r?.active || undefined);
    }
    catch (e: any) { toast(e.message, 'err'); }
  };
  const saveCell = async (it: LibItem, field: string, value: string) => {
    const r = await mepApi<{ updated: boolean; item: LibItem }>(`/api/mep/library/items/${it.id}`, { method: 'PUT', body: JSON.stringify({ [field]: value }) });
    if (r.updated) {
      const merged = (x: LibItem) => ({ ...x, ...r.item, item_type: r.item.item_type || 'Vật tư' });
      setData(d => d && {
        ...d, items: d.items.map(x => (x.id === it.id ? merged(x) : x)),
        sheet: d.sheet && { ...d.sheet, rows: d.sheet.rows.map(row => (row.item?.id === it.id ? { ...row, item: merged(row.item) } : row)) },
      });
      toast('Đã lưu');
    }
  };
  const deleteItem = async (it: LibItem) => {
    if (!confirm(`Xóa "${it.description}" khỏi ${view?.name}?`)) return;
    try { await mepApi(`/api/mep/library/items/${it.id}`, { method: 'DELETE' }); toast('Đã xóa mục'); setReload(x => x + 1); loadVersions(viewId!); }
    catch (e: any) { toast(e.message, 'err'); }
  };
  const exportView = () => {
    if (!viewId) return;
    mepDownload(`/api/mep/library/versions/${viewId}/export?${new URLSearchParams({ type, category: cat, q })}`).catch(e => toast(e.message, 'err'));
  };

  const view = versions.find(v => v.id === viewId);
  const allCount = data?.types.reduce((s, t) => s + t.count, 0) || 0;

  return (
    <div className="cx-page">
      <div className="cx-page-head">
        <div><div className="cx-crumb">Trang chủ › Thư viện giá</div><h1>Thư viện giá</h1></div>
        <div className="cx-actions">
          <button className="cx-btn" onClick={exportView} disabled={!viewId}><Icon name="download" /> Xuất version đang xem</button>
        </div>
      </div>

      <div className="cx-card cx-upload-bar">
        <h2>Nạp giá thư viện</h2>
        <label className="cx-kind-pick" title="Phân loại file trước khi chọn file nạp">Phân loại
          <select className="cx-select sm" value={kind} onChange={e => setKind(e.target.value)} disabled={busy}>
            {LIB_KINDS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
          </select>
        </label>
        <DropZone compact onFile={upload} hint="Vật tư, Nhân công, Ca máy · .xlsx · tối đa 25MB" />
        {file && (
          <span className="cx-file-line">
            <Icon name="excel" size={14} /> {file.name}
            {busy ? <span className="cx-muted">đang đọc, kiểm tra và lưu…</span>
              : result ? <span className="cx-ok">✓ {result.version.count ? `${result.version.count.toLocaleString('vi-VN')} mục` : 'bảng tra cứu'} → {result.version.name} · {libKindLabel(result.kind)} ({result.activated ? 'đang áp dụng' : 'đã lưu'})
                {Object.keys(result.by_type).length > 1 && <span className="cx-muted"> · {Object.entries(result.by_type).map(([k, n]) => `${k} ${n}`).join(' · ')}</span>}</span> : null}
          </span>
        )}
        {result && !busy && (
          <span className="cx-upload-info">
            {(result.changed.length + result.new.length + result.removed.length > 0) &&
              <button className="cx-link" onClick={() => setPopup('changes')}>{result.changed.length} đổi giá · {result.new.length} mới · {result.removed.length} không còn</button>}
          </span>
        )}
        {!file && <span className="cx-muted">Chọn phân loại rồi chọn file Excel – tool tự nhận diện cột, kiểm tra, lưu thành version mới (Đơn giá ĐNTC chuẩn được áp dụng ngay).</span>}
      </div>

      <div className="cx-lib-kindbar">
        {LIB_KINDS.map(([k, l]) => {
          const n = versions.filter(v => kindOf(v) === k).length;
          return <button key={k} className={tab === k ? 'on' : ''} onClick={() => setTab(k)}>{l}{n ? <span className="n">{n}</span> : null}</button>;
        })}
        <button className={`files ${tab === 'files' ? 'on' : ''}`} onClick={() => setTab('files')}><Icon name="history" size={14} /> File đã lưu <span className="n">{versions.length}</span></button>
      </div>

      {tab === 'kh' && <KhLearnPanel />}

      {tab === 'files' ? (
        <div className="cx-card cx-table-card cx-full">
          <div className="cx-card-head"><h2>File đã lưu ({versions.length})</h2><span className="cx-muted">Xem lại, đặt làm hiện hành hoặc xóa từng file thư viện đã nạp</span></div>
          <div className="cx-table-scroll">
            <table className="cx-table">
              <thead><tr><th>Tên</th><th>Phân loại</th><th>Trạng thái</th><th>Hiệu lực</th><th className="r">Số mục</th><th>Nguồn</th><th className="c">Thao tác</th></tr></thead>
              <tbody>
                {versions.map(v => (
                  <tr key={v.id} onDoubleClick={() => { setTab(kindOf(v)); setViewId(v.id); setSheetId(null); }}>
                    <td className="strong">{v.name}</td>
                    <td><span className={`cx-kind k-${kindOf(v)}`}>{libKindLabel(v.kind)}</span></td>
                    <td><Badge s={v.status} label={VERSION_LABEL[v.status]} /></td>
                    <td>{dmy(v.effective_date)}</td>
                    <td className="r">{v.item_count ? v.item_count.toLocaleString('vi-VN') : <span className="cx-muted">bảng tra cứu</span>}</td>
                    <td className="cx-muted">{v.note}</td>
                    <td className="c nowrap">
                      <button className="cx-btn sm" onClick={() => { setTab(kindOf(v)); setViewId(v.id); setSheetId(null); }}>Xem</button>{' '}
                      {v.status !== 'active' && !['kh', 'truc'].includes(kindOf(v)) && <><button className="cx-btn sm" onClick={() => activate(v.id)}>Đặt hiện hành</button>{' '}</>}
                      <button className="cx-btn sm danger" onClick={() => remove(v.id)}><Icon name="trash" size={13} /> Xóa</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!versions.length && <div className="cx-empty">Chưa có file nào. Chọn phân loại rồi chọn file Excel ở trên để nạp.</div>}
          </div>
        </div>
      ) : (
        <div className="cx-card cx-table-card cx-full">
          <div className="cx-card-head">
            <h2>{view ? <>{libKindLabel(tab)} · <select className="cx-select sm lib-ver-pick" value={view.id} onChange={e => { setViewId(Number(e.target.value)); setSheetId(null); }} title="Chọn bản đã nạp">
                {versions.filter(v => kindOf(v) === tab).map(v => <option key={v.id} value={v.id}>{v.name} · {VERSION_LABEL[v.status]}</option>)}
              </select></> : `${libKindLabel(tab)} · chưa có file`}</h2>
            {view && (
              <div className="cx-actions">
                <div className="cx-search sm"><Icon name="search" size={13} /><input placeholder="Tìm mã, tên, quy cách…" value={q} onChange={e => setQ(e.target.value)} /></div>
                {view.status !== 'active' && !['kh', 'truc'].includes(view.kind || '') && <button className="cx-btn sm" onClick={() => activate(view.id)}>{view.status === 'archived' ? 'Khôi phục' : 'Đặt làm hiện hành'}</button>}
              </div>
            )}
          </div>
          {data && data.sheets.length > 0 && (
            <div className="cx-tabs">
              {data.sheets.map(sh => <button key={sh.id} className={data.sheet?.sheet_id === sh.id ? 'on' : ''} onClick={() => setSheetId(sh.id)}>{sh.name} ({sh.count.toLocaleString('vi-VN')})</button>)}
              <span className="cx-tabs-note">Đúng cột, thứ tự dòng như file gốc · bấm ô để sửa</span>
            </div>
          )}
          {data && !data.sheets.length && (
            <div className="cx-tabs">
              <button className={!type ? 'on' : ''} onClick={() => { setType(''); setCat(''); }}>Tất cả ({allCount.toLocaleString('vi-VN')})</button>
              {data.types.map(t => <button key={t.name} className={type === t.name ? 'on' : ''} onClick={() => { setType(t.name); setCat(''); }}>{t.name} ({t.count.toLocaleString('vi-VN')})</button>)}
              <span className="cx-tabs-note">Bấm vào ô để sửa · Enter lưu · Esc hủy</span>
            </div>
          )}
          {data?.sheet && view && (
            <SheetGrid sheet={data.sheets.find(x => x.id === data.sheet!.sheet_id)!} rows={data.sheet.rows} compareName={data.compare_name}
              adding={adding} vid={view.id} filtered={!!(q || type || cat)}
              onAdded={ok => { setAdding(false); if (ok) { setReload(x => x + 1); loadVersions(view.id); } }}
              onSave={saveCell} onDelete={deleteItem} />
          )}
          {data && !data.sheet && <div className="cx-table-scroll">
            <table className="cx-table cx-lib-table">
              <thead>
                <tr>
                  <th className="w-type">Loại</th><th>Mã</th><th className="w-name">Tên vật tư / công việc</th><th>Quy cách / Hãng</th><th className="c">ĐVT</th>
                  <th className="r">Đơn giá VT</th><th className="r">Đơn giá NC</th><th className="r">Đơn giá máy</th><th className="c">Hao hụt</th><th className="r">Đơn giá tổng</th>
                  <th className="r">{data?.compare_name ? `So với ${data.compare_name}` : 'So sánh'}</th><th />
                </tr>
              </thead>
              <tbody>
                {adding && view && <AddRow vid={view.id} defaultType={type || 'Vật tư'} onDone={ok => { setAdding(false); if (ok) { setReload(x => x + 1); loadVersions(view.id); } }} />}
                {data?.items.map(it => (
                  <tr key={it.id}>
                    <td><EditCell value={it.item_type} options={TYPES} onSave={v => saveCell(it, 'item_type', v)} display={<span className={`cx-type t-${TYPES.indexOf(it.item_type)}`}>{it.item_type}</span>} /></td>
                    <td className="mono"><EditCell value={it.code} onSave={v => saveCell(it, 'code', v)} /></td>
                    <td className="strong"><EditCell value={it.description} onSave={v => saveCell(it, 'description', v)} /></td>
                    <td className="cx-muted"><EditCell value={it.specification} onSave={v => saveCell(it, 'specification', v)} display={[it.specification, it.brand].filter(Boolean).join(' · ') || undefined} /></td>
                    <td className="c"><EditCell value={it.unit} onSave={v => saveCell(it, 'unit', v)} /></td>
                    <td className="r"><EditCell numeric value={it.material_price} display={money(it.material_price) || undefined} onSave={v => saveCell(it, 'material_price', v)} /></td>
                    <td className="r"><EditCell numeric value={it.labor_price} display={money(it.labor_price) || undefined} onSave={v => saveCell(it, 'labor_price', v)} /></td>
                    <td className="r"><EditCell numeric value={it.machine_price} display={money(it.machine_price) || undefined} onSave={v => saveCell(it, 'machine_price', v)} /></td>
                    <td className="c"><EditCell numeric value={it.waste_percent} display={it.waste_percent !== null && it.waste_percent !== undefined ? `${num(it.waste_percent)}%` : undefined} onSave={v => saveCell(it, 'waste_percent', v)} /></td>
                    <td className="r strong"><EditCell numeric value={it.total_price} display={money(it.total_price)} onSave={v => saveCell(it, 'total_price', v)} /></td>
                    <td className="r">
                      {it.compare_percent === null
                        ? (it.is_new ? <span className="cx-badge info">Mới</span> : <span className="cx-muted">—</span>)
                        : it.compare_percent === 0 ? '0,0%' : <span className={it.compare_percent > 0 ? 'up' : 'down'}>{pctTxt(it.compare_percent)}</span>}
                    </td>
                    <td className="c"><button className="cx-icon-btn sm" title="Xóa mục" onClick={() => deleteItem(it)}><Icon name="trash" size={14} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
            {data && !data.items.length && !adding && <div className="cx-empty">Không có mục nào{q || type || cat ? ' khớp bộ lọc' : ''}.</div>}
          </div>}
          {!view && <div className="cx-empty">Chưa có file {libKindLabel(tab)}. Chọn file Excel ở trên để nạp (phân loại đã chọn sẵn theo tab này).</div>}
        </div>
      )}

      {showCompare && <CompareModal versions={versions} defaultTarget={viewId} onClose={() => setShowCompare(false)} />}
      {showHistory && view && <HistoryModal version={view} onClose={() => setShowHistory(false)} />}
      {popup && result && <ImportPopup kind={popup} r={result} onClose={() => setPopup('')} />}
    </div>
  );
}

function AddRow({ vid, defaultType, onDone, sheetId, span }: { vid: number; defaultType: string; onDone: (ok: boolean) => void; sheetId?: number; span?: number }) {
  const [f, setF] = useState<Record<string, any>>({ item_type: defaultType, ...(sheetId ? { sheet_id: sheetId } : {}) });
  const [busy, setBusy] = useState(false);
  const set = (k: string) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => setF({ ...f, [k]: e.target.value });
  const save = async () => {
    setBusy(true);
    try { await mepApi(`/api/mep/library/versions/${vid}/items`, { method: 'POST', body: JSON.stringify(f) }); toast('Đã thêm mục'); onDone(true); }
    catch (e: any) { toast(e.message, 'err'); } finally { setBusy(false); }
  };
  const key = (e: React.KeyboardEvent) => { if (e.key === 'Enter') save(); else if (e.key === 'Escape') onDone(false); };
  if (span) {
    return (
      <tr className="cx-add-row" onKeyDown={key}>
        <td colSpan={span}>
          <div className="cx-add-form">
            <b>Thêm vào cuối sheet:</b>
            <select className="cx-edit-input" value={f.item_type} onChange={set('item_type')}>{TYPES.map(t => <option key={t}>{t}</option>)}</select>
            <input className="cx-edit-input" placeholder="Mã" onChange={set('code')} />
            <input className="cx-edit-input wide" placeholder="Tên *" autoFocus onChange={set('description')} />
            <input className="cx-edit-input" placeholder="ĐVT *" onChange={set('unit')} />
            <input className="cx-edit-input r" placeholder="Đơn giá VT" onChange={set('material_price')} />
            <input className="cx-edit-input r" placeholder="Đơn giá NC" onChange={set('labor_price')} />
            <input className="cx-edit-input r" placeholder="Đơn giá máy" onChange={set('machine_price')} />
            <input className="cx-edit-input r" placeholder="Tổng (tự cộng)" onChange={set('total_price')} />
            <button className="cx-btn sm primary" disabled={busy} onClick={save}>Lưu</button>
            <button className="cx-btn sm" onClick={() => onDone(false)}>Hủy</button>
          </div>
        </td>
      </tr>
    );
  }
  return (
    <tr className="cx-add-row" onKeyDown={key}>
      <td><select className="cx-edit-input" value={f.item_type} onChange={set('item_type')}>{TYPES.map(t => <option key={t}>{t}</option>)}</select></td>
      <td><input className="cx-edit-input" placeholder="Mã" onChange={set('code')} /></td>
      <td><input className="cx-edit-input" placeholder="Tên *" autoFocus onChange={set('description')} /></td>
      <td><input className="cx-edit-input" placeholder="Quy cách" onChange={set('specification')} /></td>
      <td><input className="cx-edit-input" placeholder="ĐVT *" onChange={set('unit')} /></td>
      <td><input className="cx-edit-input r" placeholder="VT" onChange={set('material_price')} /></td>
      <td><input className="cx-edit-input r" placeholder="NC" onChange={set('labor_price')} /></td>
      <td><input className="cx-edit-input r" placeholder="Máy" onChange={set('machine_price')} /></td>
      <td><input className="cx-edit-input r" placeholder="%" onChange={set('waste_percent')} /></td>
      <td><input className="cx-edit-input r" placeholder="tự cộng" onChange={set('total_price')} /></td>
      <td colSpan={2} className="nowrap">
        <button className="cx-btn sm primary" disabled={busy} onClick={save}>Lưu</button>{' '}
        <button className="cx-btn sm" onClick={() => onDone(false)}>Hủy</button>
      </td>
    </tr>
  );
}

function CompareModal({ versions, defaultTarget, onClose }: { versions: Version[]; defaultTarget: number | null; onClose: () => void }) {
  const sorted = [...versions].sort((a, b) => (a.effective_date || '').localeCompare(b.effective_date || '') || a.id - b.id);
  const target0 = defaultTarget ?? sorted[sorted.length - 1]?.id;
  const idx = sorted.findIndex(v => v.id === target0);
  const [base, setBase] = useState<number>(idx > 0 ? sorted[idx - 1].id : sorted[idx + 1]?.id);
  const [target, setTarget] = useState<number>(target0);
  const [res, setRes] = useState<Compare | null>(null);
  const [tab, setTab] = useState<'changed' | 'added' | 'removed'>('changed');
  const [q, setQ] = useState('');
  useEffect(() => {
    if (!base || !target || base === target) { setRes(null); return; }
    mepApi<Compare>(`/api/mep/library/compare?base=${base}&target=${target}`).then(r => {
      setRes(r);
      setTab(r.summary.changed ? 'changed' : r.summary.added ? 'added' : 'removed');
    }).catch(e => toast(e.message, 'err'));
  }, [base, target]);
  const rows = useMemo(() => {
    const list = res ? res[tab] : [];
    if (!q.trim()) return list;
    const s = q.toLowerCase();
    return list.filter(r => `${r.code} ${r.description}`.toLowerCase().includes(s));
  }, [res, tab, q]);
  return (
    <Modal title="So sánh hai version" onClose={onClose} width={1080}
      footer={<>
        <button className="cx-btn" disabled={!res} onClick={() => mepDownload(`/api/mep/library/compare/export?base=${base}&target=${target}`).catch(e => toast(e.message, 'err'))}><Icon name="download" /> Xuất Excel</button>
        <button className="cx-btn primary" onClick={onClose}>Đóng</button>
      </>}>
      <div className="cx-cmp-head">
        <Field label="Version gốc (cũ)">
          <select value={base} onChange={e => setBase(Number(e.target.value))}>{versions.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}</select>
        </Field>
        <span className="cx-cmp-arrow">→</span>
        <Field label="Version so sánh (mới)">
          <select value={target} onChange={e => setTarget(Number(e.target.value))}>{versions.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}</select>
        </Field>
        {res && (
          <div className="cx-cmp-sum">
            <span className="cx-badge info">+ {res.summary.added} thêm mới</span>
            <span className="cx-badge warn">± {res.summary.changed} thay đổi</span>
            <span className="cx-badge bad">− {res.summary.removed} bị xóa</span>
            <span className="cx-badge muted">{res.summary.unchanged} không đổi</span>
          </div>
        )}
      </div>
      {base === target && <div className="cx-empty sm">Chọn hai version khác nhau.</div>}
      {res && (
        <>
          <div className="cx-tabs flush">
            <button className={tab === 'changed' ? 'on' : ''} onClick={() => setTab('changed')}>Thay đổi ({res.summary.changed})</button>
            <button className={tab === 'added' ? 'on' : ''} onClick={() => setTab('added')}>Thêm mới ({res.summary.added})</button>
            <button className={tab === 'removed' ? 'on' : ''} onClick={() => setTab('removed')}>Bị xóa ({res.summary.removed})</button>
            <div className="cx-search sm" style={{ marginLeft: 'auto' }}><Icon name="search" size={13} /><input placeholder="Tìm…" value={q} onChange={e => setQ(e.target.value)} /></div>
          </div>
          <div className="cx-table-scroll short">
            <table className="cx-table">
              <thead>
                <tr>
                  <th>Loại</th><th>Mã</th><th>Tên</th><th className="c">ĐVT</th>
                  {tab === 'changed' ? <><th>Thay đổi</th><th className="r">Đơn giá tổng {res.base.name}</th><th className="r">Đơn giá tổng {res.target.name}</th><th className="r">Chênh lệch</th></>
                    : <><th className="r">Đơn giá VT</th><th className="r">Đơn giá NC</th><th className="r">Đơn giá máy</th><th className="r">Đơn giá tổng</th></>}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => (
                  <tr key={i} className={`cmp-${tab}`}>
                    <td>{r.item_type}</td><td className="mono">{r.code}</td><td className="strong">{r.description}</td><td className="c">{r.unit}</td>
                    {tab === 'changed' ? (
                      <>
                        <td className="cx-muted">{r.field_labels?.map(l => {
                          const k = Object.keys(r.old || {}).find(f => FIELD_NAME[f] === l);
                          const ov = k ? r.old![k] : null; const nv = k ? (r as any)[k] : null;
                          return <div key={l}><b>{l}</b>: {fmt(ov)} → {fmt(nv)}</div>;
                        })}</td>
                        <td className="r">{money(r.old_total)}</td><td className="r strong">{money(r.total_price)}</td>
                        <td className="r">{r.diff_percent ? <span className={r.diff_percent > 0 ? 'up' : 'down'}>{pctTxt(r.diff_percent)}</span> : '0,0%'}</td>
                      </>
                    ) : (
                      <><td className="r">{money(r.material_price)}</td><td className="r">{money(r.labor_price)}</td><td className="r">{money(r.machine_price)}</td><td className="r strong">{money(r.total_price)}</td></>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
            {!rows.length && <div className="cx-empty sm">Không có dòng nào.</div>}
          </div>
        </>
      )}
    </Modal>
  );
}
const FIELD_NAME: Record<string, string> = {
  description: 'Tên', specification: 'Quy cách', unit: 'ĐVT', material_price: 'Đơn giá VT', labor_price: 'Đơn giá NC',
  machine_price: 'Đơn giá máy', waste_percent: 'Hao hụt', total_price: 'Đơn giá tổng',
};
const fmt = (v: any) => (v === null || v === undefined || v === '' ? '—' : typeof v === 'number' ? money(v) : String(v));

function HistoryModal({ version, onClose }: { version: Version; onClose: () => void }) {
  const [rows, setRows] = useState<Hist[] | null>(null);
  useEffect(() => { mepApi<Hist[]>(`/api/mep/library/versions/${version.id}/history`).then(setRows).catch(e => toast(e.message, 'err')); }, [version.id]);
  const label: Record<string, string> = { import: 'Nạp file', edit: 'Sửa', add: 'Thêm', delete: 'Xóa' };
  const time = (s: string) => new Date(s).toLocaleString('vi-VN', { hour12: false });
  return (
    <Modal title={`Lịch sử sửa · ${version.name}`} onClose={onClose} width={900}>
      <div className="cx-table-scroll short">
        <table className="cx-table">
          <thead><tr><th>Thời gian</th><th>Người sửa</th><th>Thao tác</th><th>Mục</th><th>Trường</th><th className="r">Giá trị cũ</th><th className="r">Giá trị mới</th></tr></thead>
          <tbody>{rows?.map(h => (
            <tr key={h.id}>
              <td className="nowrap">{time(h.changed_at)}</td><td>{h.changed_by || '—'}</td><td>{label[h.action] || h.action}</td>
              <td>{h.item_label || '—'}</td><td>{h.field_label || ''}</td>
              <td className="r cx-muted">{numOrText(h.old_value)}</td><td className="r strong">{numOrText(h.new_value)}</td>
            </tr>
          ))}</tbody>
        </table>
        {rows && !rows.length && <div className="cx-empty sm">Chưa có chỉnh sửa.</div>}
      </div>
    </Modal>
  );
}
const numOrText = (v: string | null) => (v === null || v === '' ? '—' : /^-?\d+(\.\d+)?$/.test(v) ? money(Number(v)) : v);

function ImportPopup({ kind, r, onClose }: { kind: 'errors' | 'warnings' | 'changes'; r: ImportResult; onClose: () => void }) {
  const list = kind === 'errors' ? r.errors : r.warnings;
  return (
    <Modal title={kind === 'errors' ? `Dòng lỗi bị bỏ qua (${r.errors.length})` : kind === 'warnings' ? `Cảnh báo (${r.warnings.length})` : `${r.filename} so với version áp dụng trước đó`} onClose={onClose} width={880}>
      <div className="cx-table-scroll short">
        {kind !== 'changes' ? (
          <table className="cx-table"><thead><tr><th>Sheet</th><th className="c">Dòng</th><th>Mã</th><th>Tên</th><th>{kind === 'errors' ? 'Lỗi' : 'Cảnh báo'}</th></tr></thead>
            <tbody>{list.map(e => <tr key={`${e.sheet}${e.row}`}><td>{e.sheet}</td><td className="c">{e.row}</td><td className="mono">{e.code}</td><td>{e.description}</td><td className={kind === 'errors' ? 'neg' : 'pos'}>{e.problems.join('; ')}</td></tr>)}</tbody></table>
        ) : (
          <table className="cx-table"><thead><tr><th>Mã</th><th>Tên</th><th className="r">Giá cũ</th><th className="r">Giá mới</th><th className="r">Thay đổi</th></tr></thead>
            <tbody>
              {r.changed.map(c => { const p = c.old ? ((c.new! - c.old) / c.old) * 100 : 0; return <tr key={'c' + c.code + c.description}><td className="mono">{c.code}</td><td>{c.description}</td><td className="r">{money(c.old)}</td><td className="r">{money(c.new)}</td><td className="r"><span className={p > 0 ? 'up' : 'down'}>{pctTxt(p)}</span></td></tr>; })}
              {r.new.map(c => <tr key={'n' + c.code + c.description}><td className="mono">{c.code}</td><td>{c.description}</td><td className="r cx-muted">—</td><td className="r">{money(c.new)}</td><td className="r"><span className="cx-badge info">Mới</span></td></tr>)}
              {r.removed.map(c => <tr key={'r' + c.code + c.description}><td className="mono">{c.code}</td><td>{c.description}</td><td className="r">{money(c.old)}</td><td className="r cx-muted">—</td><td className="r"><span className="cx-badge bad">Không còn</span></td></tr>)}
            </tbody></table>
        )}
      </div>
    </Modal>
  );
}

const FIELD_OF_COL = (m: Record<string, number>) => Object.fromEntries(Object.entries(m).map(([f, c]) => [c, f]));
const EDITABLE = new Set(['code', 'description', 'specification', 'brand', 'unit', 'material_price', 'labor_price', 'machine_price', 'waste_percent', 'total_price', 'category', 'item_type']);
const NUMERIC = new Set(['material_price', 'labor_price', 'machine_price', 'waste_percent', 'total_price']);
const showCell = (v: any) => (v === null || v === undefined ? '' : typeof v === 'number' ? num(v) : String(v));

function SheetGrid({ sheet, rows, compareName, adding, vid, filtered, onAdded, onSave, onDelete }: {
  sheet: Sheet; rows: SRow[]; compareName: string | null; adding: boolean; vid: number; filtered: boolean;
  onAdded: (ok: boolean) => void; onSave: (it: LibItem, field: string, v: string) => Promise<void>; onDelete: (it: LibItem) => void;
}) {
  const fieldOf = FIELD_OF_COL(sheet.mapping);
  const raw = !Object.keys(sheet.mapping || {}).length;  // sheet giữ nguyên như file (không lấy giá)
  const hasType = true;  // không thêm cột Loại riêng (chỉ hiện nếu file có cột đó)
  const cols = sheet.headers.map((h, i) => ({ h, col: sheet.first_col + i, field: fieldOf[sheet.first_col + i] as string | undefined }));
  const span = cols.length + (raw ? 1 : 2);
  const cls = (f?: string) => (f === 'description' ? 'w-name' : f && NUMERIC.has(f) ? 'r' : f === 'unit' || f === 'stt' ? 'c' : f === 'code' ? 'nowrap' : '');
  // Kéo mép phải tiêu đề để đổi độ rộng cột (nhớ theo tên sheet trên máy này; nhấp đúp mép = trả về tự động)
  const wKey = `cx-libcol-${sheet.name}`;
  const [colW, setColW] = useState<number[] | null>(null);
  useEffect(() => { try { const v = JSON.parse(localStorage.getItem(wKey) || 'null'); setColW(Array.isArray(v) && v.length === span ? v : null); } catch { setColW(null); } }, [wKey, span]);
  const tblRef = useRef<HTMLTableElement>(null);
  const startResize = (e: React.MouseEvent, ci: number) => {
    e.preventDefault(); e.stopPropagation();
    const t = tblRef.current; if (!t) return;
    const ths = Array.from(t.querySelectorAll<HTMLTableCellElement>('thead tr:first-child th'));
    const base = colW || ths.map(th => Math.round(th.getBoundingClientRect().width));
    const x0 = e.clientX, w0 = base[ci];
    let cur = base.slice();
    document.body.classList.add('cx-resizing-col');
    const mv = (ev: MouseEvent) => { cur = base.slice(); cur[ci] = Math.max(30, Math.round(w0 + ev.clientX - x0)); setColW(cur); };
    const up = () => {
      window.removeEventListener('mousemove', mv); window.removeEventListener('mouseup', up); document.body.classList.remove('cx-resizing-col');
      try { localStorage.setItem(wKey, JSON.stringify(cur)); } catch { /* bỏ qua */ }
    };
    window.addEventListener('mousemove', mv); window.addEventListener('mouseup', up);
  };
  const resetW = (e: React.MouseEvent) => { e.stopPropagation(); setColW(null); try { localStorage.removeItem(wKey); } catch { /* bỏ qua */ } };
  const rz = (ci: number) => <span className="lib-rz" title="Kéo để đổi độ rộng cột · nhấp đúp để trả về tự động" onMouseDown={e => startResize(e, ci)} onDoubleClick={resetW} />;
  const titles = rows.filter(r => r.kind === 'title');
  // Cấp dòng nhóm theo STT (như file): A, B… → cấp 1; I, II… → cấp 2; I.1, 1.1… → cấp 3 (tô màu đậm nhạt)
  const isStt = (h: string) => /(^|\s)(stt|no\.?)\s*$/i.test((h || '').trim()) || /^\s*(stt|no\.?)(\s|$)/i.test((h || '').trim());
  const sttIdx = Math.max(0, cols.findIndex(c => isStt(c.h || '')));
  // Canh giữa các cột STT, ĐV / ĐVT, KL (khối lượng)
  const centerCols = cols.map((c, i) => [c, i] as const).filter(([c]) => (/^\s*(stt|đv|dv|đvt|dvt|đơn vị|kl|khối lượng|unit|quantity)\b/i.test((c.h || '').normalize('NFC')) && !/price/i.test(c.h || '')) || isStt(c.h || '') || c.field === 'unit').map(([, i]) => i);
  const lvl = (r: SRow) => { const t = String(r.cells[sttIdx] ?? '').trim(); return /^[A-Z]$/.test(t) ? 1 : /^[IVXL]+$/.test(t) ? 2 : t ? 3 : 0; };
  return (
    <div className="cx-table-scroll lib-sheet-scroll">
      {centerCols.length > 0 && <style>{centerCols.map(i => `.cx-sheet-grid tbody tr:not(.k-note) > td:nth-child(${i + 2}), .cx-sheet-grid thead th:nth-child(${i + 2})`).join(', ') + ' { text-align: center !important; }'}</style>}
      {titles.length > 0 && <div className="lib-titles">{titles.map(t => <div key={t.id}>{String(t.cells[0] ?? '')}</div>)}</div>}
      <table ref={tblRef} className={`cx-table cx-lib-table cx-sheet-grid ${colW ? 'fixedw' : ''}`} style={colW ? { width: colW.reduce((a, b) => a + b, 0) } : undefined}>
        {colW && <colgroup>{colW.map((w, i) => <col key={i} style={{ width: w }} />)}</colgroup>}
        <thead>
          <tr>
            <th className="c rowno" title="Dòng trong file Excel gốc">Dòng{rz(0)}</th>
            {cols.map((c, i) => <th key={c.col} className={`${cls(c.field)} ${i === sttIdx && isStt(c.h || '') ? 'stt-col' : ''}`} title={c.field ? `Tool hiểu là: ${c.field}` : 'Cột giữ nguyên, không dùng để tính'}>{c.h || <span className="cx-muted">(trống)</span>}{rz(i + 1)}</th>)}
            {!hasType && <th className="w-type">Loại</th>}
            {!raw && <th />}
          </tr>
        </thead>
        <tbody>
          {rows.map(r => {
            // Dòng ghi chú (vd "Ghi chú: Đơn giá = % khấu hao…") trong sheet đơn giá: ẩn trên web, file gốc vẫn giữ
            if (!raw && r.kind !== 'item' && /^\s*ghi\s*ch[uú]/i.test(String(r.cells.find(x => x !== null && x !== undefined && String(x).trim() !== '') ?? '').normalize('NFC'))) return null;
            if (r.kind === 'title') return null;
            // Dòng ghi chú / diễn giải dài (1 ô chữ, thường gộp ô trong file): trải hết bề ngang, xuống dòng như file để dễ đọc
            const filled = r.cells.map((v, i) => [v, i] as const).filter(([v]) => v !== null && v !== undefined && String(v).trim() !== '');
            if (raw && filled.length === 1 && typeof filled[0][0] === 'string' && (filled[0][0].length > 60 || /^\s*ghi\s*ch/i.test(filled[0][0].normalize('NFC')))) {
              const at = filled[0][1];
              return (
                <tr key={r.id} className="k-note">
                  <td className="c rowno">{r.row ?? ''}</td>
                  {cols.slice(0, at).map(c => <td key={c.col} />)}
                  <td colSpan={cols.length - at} className="lib-note">{String(filled[0][0]).trim()}</td>
                </tr>
              );
            }
            if (r.kind !== 'item' || !r.item) {
              return (
                <tr key={r.id} className={r.kind === 'error' || r.kind === 'raw' ? `k-plain ${raw && lvl(r) && lvl(r) < 3 && !/^\d/.test(String(r.cells[sttIdx] ?? '')) ? `k-group lv${lvl(r)}` : ''}` : `k-${r.kind} lv${lvl(r)}`} title={r.kind === 'deleted' ? 'Đã xóa trên web' : undefined}>
                  <td className="c rowno">{r.row ?? ''}</td>
                  {cols.map((c, i) => <td key={c.col} className={cls(c.field) || (typeof r.cells[i] === 'number' || /^-?[\d.,]+\s*%?$/.test(String(r.cells[i] ?? '').trim()) && i !== sttIdx ? 'r' : '')}>{showCell(r.cells[i])}</td>)}
                  {span - cols.length - 1 > 0 && <td colSpan={span - cols.length - 1} className="cx-muted">{r.kind === 'deleted' ? 'đã xóa' : ''}</td>}
                </tr>
              );
            }
            const it = r.item;
            return (
              <tr key={r.id} className={r.row === null ? 'k-new' : ''}>
                <td className="c rowno">{r.row ?? 'mới'}</td>
                {cols.map((c, i) => {
                  const f = c.field;
                  if (!f || !EDITABLE.has(f)) return <td key={c.col} className={cls(f)}>{showCell(r.cells[i])}</td>;
                  const v = (it as any)[f];
                  if (f === 'item_type') return <td key={c.col}><EditCell value={v} options={TYPES} onSave={x => onSave(it, f, x)} /></td>;
                  const disp = NUMERIC.has(f) ? (f === 'waste_percent' && v !== null && v !== undefined ? `${num(v)}%` : money(v) || undefined) : undefined;
                  return <td key={c.col} className={`${cls(f)} ${f === 'total_price' ? 'strong' : ''}`}><EditCell numeric={NUMERIC.has(f)} value={v} display={disp} onSave={x => onSave(it, f, x)} /></td>;
                })}
                {!hasType && <td><EditCell value={it.item_type} options={TYPES} onSave={x => onSave(it, 'item_type', x)} display={<span className={`cx-type t-${TYPES.indexOf(it.item_type)}`}>{it.item_type}</span>} /></td>}
                <td className="c"><button className="cx-icon-btn sm" title="Xóa mục" onClick={() => onDelete(it)}><Icon name="trash" size={14} /></button></td>
              </tr>
            );
          })}
          {adding && <AddRow vid={vid} defaultType="Vật tư" sheetId={sheet.id} span={span} onDone={onAdded} />}
        </tbody>
      </table>
      {!rows.length && !adding && <div className="cx-empty">{filtered ? 'Không có mục nào khớp bộ lọc trong sheet này.' : 'Sheet trống.'}</div>}
    </div>
  );
}


// Vật tư khấu hao đã học: tên nào có thể khấu hao dưới 100% / tên nào mặc định 100%
function KhLearnPanel() {
  const [d, setD] = useState<{ lt100: string[]; full100: string[] } | null>(null);
  const [busy, setBusy] = useState(false);
  const [q, setQ] = useState('');
  const [open, setOpen] = useState(true);
  const fileRef = useRef<HTMLInputElement>(null);
  const load = () => mepApi<{ lt100: string[]; full100: string[] }>('/api/mep/library/depreciation').then(setD).catch(e => toast(e.message, 'err'));
  useEffect(() => { load(); }, []);
  const learn = async (f: File) => {
    setBusy(true);
    const fd = new FormData(); fd.append('file', f);
    try {
      const r = await mepApi<any>('/api/mep/library/depreciation/learn', { method: 'POST', body: fd });
      toast(`Đã học từ ${f.name}: thêm ${r.added_lt100} tên có khấu hao < 100%, ${r.added_100} tên mặc định 100%`);
      load();
    } catch (e: any) { toast(e.message, 'err'); } finally { setBusy(false); }
  };
  const move = async (name: string, kh: 'lt100' | '100') => {
    try { setD(await mepApi(`/api/mep/library/depreciation/set`, { method: 'POST', body: JSON.stringify({ name, kh }) })); } catch (e: any) { toast(e.message, 'err'); }
  };
  const nq = q.trim().toLowerCase();
  const f = (xs: string[]) => xs.filter(x => !nq || x.toLowerCase().includes(nq));
  return (
    <div className="cx-card kh-learn">
      <div className="cx-card-head">
        <h2 style={{ cursor: 'pointer' }} onClick={() => setOpen(!open)}>{open ? '▾' : '▸'} Vật tư khấu hao đã học</h2>
        <div className="cx-actions">
          <div className="cx-search sm"><Icon name="search" size={13} /><input placeholder="Tìm tên vật tư…" value={q} onChange={e => setQ(e.target.value)} /></div>
          <button className="cx-btn sm" disabled={busy} onClick={() => fileRef.current?.click()} title="Chọn hồ sơ BOQ có cột % KH VẬT TƯ để tool học thêm">
            <Icon name="upload" size={14} /> {busy ? 'Đang học…' : 'Học từ hồ sơ BOQ'}</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xlsm" hidden onChange={e => { const x = e.target.files?.[0]; if (x) learn(x); e.target.value = ''; }} />
        </div>
      </div>
      {open && d && (
        <div className="kh-cols">
          {([['lt100', 'Có thể khấu hao dưới 100%', d.lt100, '100', 'Chuyển sang mặc định 100%'], ['100', 'Mặc định khấu hao 100%', d.full100, 'lt100', 'Chuyển sang có thể < 100%']] as const).map(([k, title, xs, to, tip]) => (
            <div key={k} className={`kh-col k-${k}`}>
              <div className="kh-col-h">{title} <span className="n">{xs.length}</span></div>
              <div className="kh-list">
                {f([...xs]).map(x => <div key={x} className="kh-item"><span>{x}</span><button title={tip} onClick={() => move(x, to)}>⇄</button></div>)}
                {!f([...xs]).length && <div className="cx-muted sm">Không có</div>}
              </div>
            </div>
          ))}
        </div>
      )}
      {open && <div className="cx-muted kh-note">Tool chỉ lưu tên vật tư và loại (dưới 100% / 100%), không lưu số %. Vật tư chưa học: suy theo nhóm trong bảng KH VẬT TƯ (vd cáp lớn, tủ điện, đèn, bơm có thể &lt; 100%; cáp nhỏ, ống, van, phụ kiện mặc định 100%).</div>}
    </div>
  );
}

