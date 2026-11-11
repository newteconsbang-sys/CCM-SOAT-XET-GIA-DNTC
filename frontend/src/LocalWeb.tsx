import React, { useEffect, useState } from 'react';
import { mepApi } from './mepApi';
import { ReviewList } from './ccm/ReviewList';
import { PackageDetail } from './ccm/PackageDetail';
import { LibraryVersions } from './ccm/LibraryVersions';
import { Icon, Toaster } from './ccm/ui';
import './mepStyles.css';
import './ccm/theme.css';
import './ccm/theme-v2.css';
import logoUrl from './ccm/newtecons-logo.png';

type Route = { page: 'review' | 'package' | 'library'; id?: number };

function readRoute(): Route {
  const h = location.hash.replace(/^#\/?/, '');
  const [page, id] = h.split('/');
  if (page === 'package' && Number(id)) return { page: 'package', id: Number(id) };
  if (page === 'library') return { page: 'library' };
  return { page: 'review' };
}

function initials(name: string) {
  const w = name.replace(/^(KS|ThS|TS|Kts)\.?\s+/i, '').trim().split(/\s+/).filter(Boolean);
  return ((w[w.length - 2]?.[0] || '') + (w[w.length - 1]?.[0] || '')).toUpperCase() || 'CC';
}

export function LocalWeb() {
  const [route, setRoute] = useState<Route>(readRoute);
  const [ready, setReady] = useState<boolean | null>(null);
  const [user, setUser] = useState(() => { try { return localStorage.getItem('cx-user') || 'Người soát xét'; } catch { return 'Người soát xét'; } });

  useEffect(() => {
    const h = () => setRoute(readRoute());
    window.addEventListener('hashchange', h);
    mepApi('/api/health').then(() => setReady(true)).catch(() => setReady(false));
    return () => window.removeEventListener('hashchange', h);
  }, []);
  const go = (r: string) => { location.hash = `#/${r}`; };
  const renameUser = () => {
    const n = prompt('Tên hiển thị của bạn (VD: KS. Nguyễn Văn An):', user);
    if (n && n.trim()) { setUser(n.trim()); try { localStorage.setItem('cx-user', n.trim()); } catch { /* bỏ qua */ } }
  };

  const nav = (id: Route['page'], label: string, icon: string, also: Route['page'][] = []) => (
    <button className={`cx-nav ${route.page === id || also.includes(route.page) ? 'on' : ''}`} onClick={() => go(id)}>
      <Icon name={icon} size={18} /> <span>{label}</span>
    </button>
  );

  return (
    <div className="cx-app">
      <aside className="cx-sidebar">
        <div className="cx-brand">
          <div className="cx-logo-card" title="CCM-AI – Soát xét giá điện nước"><img src={logoUrl} alt="NEWTECONS – Build on trust" /></div>
        </div>
        <div className="cx-nav-group">CHỨC NĂNG</div>
        {nav('review', 'Soát xét giá', 'review', ['package'])}
        {nav('library', 'Thư viện giá', 'library')}
        <button className="cx-side-user" onClick={renameUser} title="Tên người soát xét (ghi vào lịch sử chỉnh sửa) – bấm để đổi">
          <span className="cx-avatar">{initials(user)}</span><span>{user}</span>
        </button>
        <div className="cx-side-foot">
          <span className={`cx-dot ${ready ? 'on' : ready === false ? 'off' : ''}`} />
          {ready ? 'Đã kết nối dịch vụ' : ready === false ? 'Chưa kết nối – mở Khởi động CCM-AI' : 'Đang kết nối…'}
          <div>Phiên bản 0.9.101 · © 2026</div>
        </div>
      </aside>

      <div className="cx-main">
        <div className="cx-content">
          {route.page === 'review' && <ReviewList onOpen={id => go(`package/${id}`)} />}
          {route.page === 'package' && route.id && <PackageDetail key={route.id} id={route.id} onExit={() => go('review')} />}
          {route.page === 'library' && <LibraryVersions />}
        </div>
      </div>
      <Toaster />
    </div>
  );
}
