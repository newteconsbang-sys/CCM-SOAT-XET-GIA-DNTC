/**
 * API client for CCM-AI MEP Price Review System.
 */
let sessionToken = '';

export async function getSessionToken(): Promise<string> {
  if (sessionToken) return sessionToken;
  try {
    const res = await fetch('/api/session');
    if (res.ok) {
      const data = await res.json();
      sessionToken = data.token || '';
    }
  } catch (err) {
    console.warn('Failed to fetch session token:', err);
  }
  return sessionToken;
}

export async function mepApi<T = any>(endpoint: string, options: RequestInit = {}): Promise<T> {
  const token = await getSessionToken();
  const headers = new Headers(options.headers || {});
  if (token) {
    headers.set('X-CCM-Token', token);
  }
  try {
    const user = localStorage.getItem('cx-user');
    if (user) headers.set('X-CCM-User', encodeURIComponent(user));
  } catch { /* trình duyệt chặn localStorage */ }
  if (!(options.body instanceof FormData) && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  const response = await fetch(endpoint, {
    ...options,
    headers,
  });

  if (response.status === 401) {
    // Retry once with new token
    sessionToken = '';
    const newToken = await getSessionToken();
    headers.set('X-CCM-Token', newToken);
    const retry = await fetch(endpoint, { ...options, headers });
    if (!retry.ok) {
      const err = await retry.json().catch(() => ({ detail: retry.statusText }));
      throw new Error(err.detail || 'Lỗi xác thực hệ thống');
    }
    return retry.json();
  }

  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: response.statusText }));
    throw new Error(err.detail || `Lỗi máy chủ (${response.status})`);
  }

  return response.json();
}


/** Tải file từ API có token phiên (thẻ <a href> thường sẽ bị 401). */
export async function mepDownload(endpoint: string, fallbackName = 'tai-ve.xlsx'): Promise<string | null> {
  const token = await getSessionToken();
  const extra: Record<string, string> = {};
  try { const u = localStorage.getItem('cx-user'); if (u) extra['X-CCM-User'] = encodeURIComponent(u); } catch { /* bỏ qua */ }
  let res = await fetch(endpoint, { headers: { 'X-CCM-Token': token, ...extra } });
  if (res.status === 401) {
    sessionToken = '';
    res = await fetch(endpoint, { headers: { 'X-CCM-Token': await getSessionToken() } });
  }
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || `Lỗi tải file (${res.status})`);
  }
  const cd = res.headers.get('Content-Disposition') || '';
  const m = /filename\*=UTF-8''([^;]+)/i.exec(cd) || /filename="?([^";]+)"?/i.exec(cd);
  const name = m ? decodeURIComponent(m[1]) : fallbackName;
  const url = URL.createObjectURL(await res.blob());
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return res.headers.get('X-CCM-Warning') || null;
}
