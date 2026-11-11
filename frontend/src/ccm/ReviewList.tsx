import React, { useEffect, useRef, useState } from 'react';
import { mepApi } from '../mepApi';
import { Badge, DropZone, Field, Icon, Modal, Pkg, Project, RowMenu, Version, money, pct, toast, PAS } from './ui';

const projectCode = (p: Project) => `DA${String(p.id).padStart(4, '0')}`;
const noAccent = (s: string) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'd').toLowerCase();

export function ReviewList({ onOpen }: { onOpen: (id: number) => void }) {
  const [projects, setProjects] = useState<Project[]>([]);
  const [versions, setVersions] = useState<Version[]>([]);
  const [pkgs, setPkgs] = useState<Pkg[]>([]);
  const [selProject, setSelProjectState] = useState<number | null>(() => {
    try { const v = Number(localStorage.getItem('cx-project')); return v || null; } catch { return null; }
  });
  const [loading, setLoading] = useState(false);
  const [showPkg, setShowPkg] = useState<null | { project_id?: number }>(null);
  const [showProject, setShowProject] = useState(false);

  const setSelProject = (id: number | null) => {
    setSelProjectState(id);
    try { id ? localStorage.setItem('cx-project', String(id)) : localStorage.removeItem('cx-project'); } catch { /* bỏ qua */ }
  };
  const loadAll = async () => {
    setLoading(true);
    try {
      const [p, v, k] = await Promise.all([mepApi<Project[]>('/api/mep/projects'), mepApi<Version[]>('/api/mep/library/versions'),
        mepApi<{ items: Pkg[] }>('/api/mep/packages')]);
      setProjects(p); setVersions(v); setPkgs(k.items);
    } catch (e: any) { toast(e.message, 'err'); } finally { setLoading(false); }
  };
  const reload = () => { loadAll(); };
  useEffect(() => { loadAll(); }, []);
  // Dự án đã chọn bị xóa thì quay về "Tất cả dự án"
  useEffect(() => { if (selProject && projects.length && !projects.some(p => p.id === selProject)) setSelProject(null); }, [projects]);

  const setStatus = async (p: Pkg, status: string) => {
    try { await mepApi(`/api/mep/packages/${p.id}`, { method: 'PUT', body: JSON.stringify({ status }) }); toast(`${p.code}: ${status}`); reload(); }
    catch (e: any) { toast(e.message, 'err'); }
  };
  const remove = async (p: Pkg) => {
    if (!confirm(`Xóa gói thầu ${p.code} – ${p.name}?`)) return;
    try { await mepApi(`/api/mep/packages/${p.id}`, { method: 'DELETE' }); toast('Đã xóa gói thầu'); reload(); }
    catch (e: any) { toast(e.message, 'err'); }
  };
  const removeProject = async (p: Project) => {
    const n = p.packages.length;
    const msg = n ? `Dự án "${p.name}" đang có ${n} gói thầu.\nXóa dự án sẽ xóa luôn ${n} gói thầu này. Tiếp tục?` : `Xóa dự án "${p.name}"?`;
    if (!confirm(msg)) return;
    try {
      await mepApi(`/api/mep/projects/${p.id}${n ? '?force=true' : ''}`, { method: 'DELETE' });
      toast(`Đã xóa dự án ${p.name}${n ? ` và ${n} gói thầu` : ''}`);
      if (selProject === p.id) setSelProject(null);
      reload();
    } catch (e: any) { toast(e.message, 'err'); }
  };
  const activeVersion = versions.find(v => v.status === 'active');
  const current = projects.find(p => p.id === selProject);
  const shown = selProject ? pkgs.filter(p => p.project_id === selProject) : pkgs;

  return (
    <div className="cx-page">
      <div className="cx-page-head cx-head-bar">
        <ProjectPicker projects={projects} value={selProject} onChange={setSelProject} onCreate={() => setShowProject(true)} onDelete={removeProject} />
        <div className="cx-actions">
          {current && <button className="cx-btn danger" onClick={() => removeProject(current)} title="Xóa dự án đang chọn"><Icon name="trash" /> Xóa dự án</button>}
          <button className="cx-btn" onClick={() => setShowProject(true)}><Icon name="plus" /> Tạo dự án</button>
          <button className="cx-btn primary" onClick={() => setShowPkg({ project_id: selProject || undefined })}><Icon name="plus" /> Tạo gói thầu</button>
        </div>
      </div>

      <div className="cx-card cx-table-card cx-full">
        <div className="cx-card-head">
          <h2>{current ? `Gói thầu · ${current.name}` : 'Gói thầu · Tất cả dự án'} ({shown.length})</h2>
          <span className="cx-muted">Đơn vị: đồng (chưa VAT)</span>
        </div>
        <div className="cx-table-scroll">
          <table className="cx-table">
            <thead>
              <tr>
                <th>Mã gói</th><th className="w-name">Tên gói thầu</th>{!current && <th>Dự án</th>}<th>Nhà thầu phụ</th>
                <th className="r">Giá NTP<br />báo</th><th className="r">Giá soát<br />xét</th><th className="r">Chênh lệch<br />(đồng)</th><th className="r">Chênh<br />lệch %</th>
                <th className="c">Đối chiếu<br />theo</th><th className="c">Trạng thái</th><th className="c">Thao tác</th>
              </tr>
            </thead>
            <tbody>
              {shown.map(p => (
                <tr key={p.id} onDoubleClick={() => onOpen(p.id)}>
                  <td className="mono">{p.code}</td>
                  <td className="strong">{p.name}</td>
                  {!current && <td>{p.project_name}</td>}
                  <td>{p.subcontractor || '—'}</td>
                  <td className="r num-full">{p.ntp_total ? money(p.ntp_total) : '—'}</td>
                  <td className="r num-full strong">{p.review_total ? money(p.review_total) : '—'}</td>
                  {(() => { const df = p.ntp_total && p.review_total ? p.review_total - p.ntp_total : null; return <>
                    <td className={`r num-full ${df && df < 0 ? 'neg' : df && df > 0 ? 'pos' : ''}`}>{df === null ? '—' : money(df)}</td>
                    <td className={`r ${p.diff_percent && p.diff_percent < 0 ? 'neg' : p.diff_percent && p.diff_percent > 0 ? 'pos' : ''}`}>{p.diff_percent === null ? '—' : pct(p.diff_percent)}</td></>; })()}
                  <td className="c" title={p.version_name}>{p.version_name?.startsWith('Đối chiếu') ? 'File đính kèm' : p.version_name || '—'}</td>
                  <td className="c"><Badge s={p.status} /></td>
                  <td className="c nowrap">
                    <button className="cx-btn sm" onClick={() => onOpen(p.id)}>Mở</button>{' '}
                    <RowMenu items={[
                      { label: 'Đã phản hồi BCHCT', onClick: () => setStatus(p, 'Đã phản hồi BCHCT'), hidden: !p.session_id || p.status === 'Đã phản hồi BCHCT' },
                      { label: 'Xong', onClick: () => setStatus(p, 'Xong'), hidden: !p.session_id || p.status === 'Xong' },
                      { label: 'Đưa về đang soát xét', onClick: () => setStatus(p, 'Đang soát xét'), hidden: !p.session_id || p.status === 'Đang soát xét' },
                      { label: 'Xóa gói thầu', onClick: () => remove(p), danger: true },
                    ]} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {!loading && shown.length === 0 && (
            <div className="cx-empty">
              {projects.length === 0 ? <>Chưa có dự án. Bấm <b>Tạo dự án</b>, sau đó <b>Tạo gói thầu</b> và nạp file báo giá NTP.</>
                : <>Dự án này chưa có gói thầu. Bấm <b>Tạo gói thầu</b> để nạp file báo giá NTP.</>}
            </div>
          )}
          {loading && <div className="cx-empty">Đang tải…</div>}
        </div>
      </div>

      {showProject && <ProjectModal onClose={() => setShowProject(false)} onDone={id => { setShowProject(false); setSelProject(id); reload(); }} />}
      {showPkg && (
        <PackageModal
          projects={projects} versions={versions} subs={[]} activeVersion={activeVersion} initialProject={showPkg.project_id}
          onClose={() => setShowPkg(null)}
          onDone={(id, open) => { setShowPkg(null); reload(); if (open) onOpen(id); }}
          onNeedProject={() => { setShowPkg(null); setShowProject(true); }}
        />
      )}
    </div>
  );
}

/** Chọn dự án kiểu ô tìm kiếm thả xuống (như CFMS): gõ để lọc theo mã hoặc tên. */
function ProjectPicker({ projects, value, onChange, onCreate, onDelete }: { projects: Project[]; value: number | null; onChange: (id: number | null) => void; onCreate: () => void; onDelete: (p: Project) => void }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const h = (e: MouseEvent) => !ref.current?.contains(e.target as Node) && setOpen(false);
    document.addEventListener('mousedown', h);
    return () => document.removeEventListener('mousedown', h);
  }, [open]);
  const cur = projects.find(p => p.id === value);
  const list = projects.filter(p => !q.trim() || noAccent(`${projectCode(p)} ${p.name}`).includes(noAccent(q.trim())));
  const pick = (id: number | null) => { onChange(id); setOpen(false); setQ(''); };
  return (
    <div className="cx-picker" ref={ref}>
      <button className={`cx-picker-btn ${open ? 'open' : ''}`} onClick={() => setOpen(o => !o)}>
        <Icon name="library" size={16} />
        <span className="cx-picker-label">{cur ? `${projectCode(cur)} - ${cur.name}` : 'Tất cả dự án'}</span>
        <span className="cx-picker-chev">▾</span>
      </button>
      {open && (
        <div className="cx-picker-pop">
          <div className="cx-picker-search"><Icon name="search" size={16} /><input autoFocus placeholder="Tìm kiếm dự án…" value={q} onChange={e => setQ(e.target.value)}
            onKeyDown={e => { if (e.key === 'Escape') setOpen(false); if (e.key === 'Enter' && list[0]) pick(list[0].id); }} /></div>
          <div className="cx-picker-list">
            {!q && <button className={!value ? 'on' : ''} onClick={() => pick(null)}>Tất cả dự án</button>}
            {list.map(p => (
              <div key={p.id} className={`cx-picker-row ${p.id === value ? 'on' : ''}`}>
                <button onClick={() => pick(p.id)}>
                  <span className="mono">{projectCode(p)}</span> - {p.name} <small>{p.packages.length} gói thầu</small>
                </button>
                <button className="cx-picker-del" title="Xóa dự án" onClick={() => { setOpen(false); onDelete(p); }}><Icon name="trash" size={14} /></button>
              </div>
            ))}
            {!list.length && <div className="cx-empty sm">Không thấy dự án “{q}”. <a className="cx-link" onClick={() => { setOpen(false); onCreate(); }}>Tạo dự án mới</a></div>}
          </div>
        </div>
      )}
    </div>
  );
}

function ProjectModal({ onClose, onDone }: { onClose: () => void; onDone: (id: number) => void }) {
  const [name, setName] = useState('');
  const [region, setRegion] = useState('xa');
  const save = async () => {
    try {
      const r = await mepApi<{ id: number }>('/api/mep/projects', { method: 'POST', body: JSON.stringify({ name, kh_region: region }) });
      toast(`Đã tạo dự án ${name}`); onDone(r.id);
    } catch (e: any) { toast(e.message, 'err'); }
  };
  return (
    <Modal title="Tạo dự án mới" onClose={onClose} width={520}
      footer={<><button className="cx-btn" onClick={onClose}>Hủy</button><button className="cx-btn primary" disabled={!name.trim()} onClick={save}>Tạo dự án</button></>}>
      <div className="cx-form">
        <Field label="Tên dự án" req span={2}><input autoFocus value={name} onChange={e => setName(e.target.value)} placeholder="VD: Chung Cư Ruby Tower" onKeyDown={e => e.key === 'Enter' && name.trim() && save()} /></Field>
        <Field label="Vị trí (tra bảng khấu hao vật tư)" req span={2}>
          <select value={region} onChange={e => setRegion(e.target.value)}><option value="xa">XA BIỂN</option><option value="gan">GẦN BIỂN</option></select>
        </Field>
      </div>
    </Modal>
  );
}

function nextCode(project?: Project): string {
  const codes = project?.packages.map(p => p.code) || [];
  const last = codes.map(c => /^(.*?)(\d+)$/.exec(c)).filter(Boolean).sort((a, b) => Number(a![2]) - Number(b![2])).pop();
  if (last) return last[1] + String(Number(last[2]) + 1).padStart(last[2].length, '0');
  if (!project) return '';
  const initials = project.name.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/gi, 'D')
    .split(/\s+/).filter(w => w && !/^(chung|cu|du|an)$/i.test(w)).map(w => w[0]).join('').toUpperCase().slice(0, 3);
  return `${initials || 'GT'}-ĐN-01`;
}

function PackageModal({ projects, versions, activeVersion, initialProject, onClose, onDone, onNeedProject }: {
  projects: Project[]; versions: Version[]; subs: string[]; activeVersion?: Version; initialProject?: number;
  onClose: () => void; onDone: (id: number, open: boolean) => void; onNeedProject: () => void;
}) {
  const [projectId, setProjectId] = useState<number | ''>(initialProject || projects[0]?.id || '');
  const [name, setName] = useState('');
  const [pa, setPa] = useState('PA2');
  const [region, setRegion] = useState('xa');
  useEffect(() => { const pr = projects.find(p => p.id === projectId) as any; setRegion(pr?.kh_region === 'gan' ? 'gan' : 'xa'); }, [projectId]);
  const [busy, setBusy] = useState(false);
  const valid = !!(projectId && name.trim());

  const submit = async () => {
    if (!valid) { toast('Nhập tên gói thầu', 'err'); return; }
    setBusy(true);
    const fd = new FormData();
    fd.append('project_id', String(projectId)); fd.append('name', name.trim()); fd.append('pa', pa); fd.append('kh_region', region);
    if (activeVersion?.id || versions[0]?.id) fd.append('version_id', String(activeVersion?.id || versions[0]?.id));
    try {
      const r = await mepApi<{ id: number; code: string; warning?: string }>('/api/mep/packages', { method: 'POST', body: fd });
      toast(`Đã tạo gói ${r.code}`);
      onDone(r.id, true);
    } catch (e: any) { toast(e.message, 'err'); } finally { setBusy(false); }
  };

  if (!projects.length) {
    return (
      <Modal title="Tạo gói thầu mới" onClose={onClose} width={520} footer={<><button className="cx-btn" onClick={onClose}>Hủy</button><button className="cx-btn primary" onClick={onNeedProject}>Tạo dự án</button></>}>
        <p>Cần có ít nhất một dự án trước khi tạo gói thầu.</p>
      </Modal>
    );
  }
  return (
    <Modal title="Tạo gói thầu mới" onClose={onClose}
      footer={<>
        <button className="cx-btn" onClick={onClose} disabled={busy}>Hủy</button>
        <button className="cx-btn primary" onClick={submit} disabled={!valid || busy}>{busy ? 'Đang tạo…' : 'Tạo gói'}</button>
      </>}>
      <div className="cx-form">
        <Field label="Dự án" req span={2}>
          <select value={projectId} onChange={e => setProjectId(Number(e.target.value))}>
            {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
        </Field>
        <Field label="Tên gói thầu" req span={2}><input autoFocus value={name} onChange={e => setName(e.target.value)} placeholder="VD: Hệ thống chống sét & nối đất" /></Field>
        <Field label="Phương án" req span={2}>
          <select value={pa} onChange={e => setPa(e.target.value)}>{PAS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
          {pa === 'PA1' && <small className="cx-hint">Dự toán: file phải có đủ các cột A → M (STT, Mô tả, Đơn vị, Khối lượng, ĐG VT, ĐG NC, % đã KH, % KH, Bắt đầu, Kết thúc, Đơn giá, Thành tiền, Ghi chú). Nhập A–D, I, J – tool tự áp giá cột E, F và %KH cột H theo thư viện giá.</small>}
        </Field>
        <Field label="Vị trí dự án (tra bảng khấu hao vật tư)" req span={2}>
          <select value={region} onChange={e => setRegion(e.target.value)}><option value="xa">XA BIỂN</option><option value="gan">GẦN BIỂN</option></select>
        </Field>
        <p className="cx-hint" style={{ gridColumn: '1 / -1', margin: 0 }}>Tạo xong sẽ mở gói – bấm <b>Nhập dữ liệu từ excel</b> để nạp file khi cần.</p>
      </div>
    </Modal>
  );
}
