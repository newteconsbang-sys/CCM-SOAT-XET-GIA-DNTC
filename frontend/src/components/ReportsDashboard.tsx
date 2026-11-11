import React, { useState, useEffect } from 'react';
import { mepApi } from '../mepApi';

export function ReportsDashboard() {
  const [sessions, setSessions] = useState<any[]>([]);
  const [selectedSessionId, setSelectedSessionId] = useState<string>('');
  const [sessionData, setSessionData] = useState<any>(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const list = await mepApi('/api/mep/history/sessions');
        setSessions(list || []);
        if (list && list.length > 0) {
          setSelectedSessionId(list[0].id);
        }
      } catch (err: any) {
        console.error(err);
      }
    })();
  }, []);

  useEffect(() => {
    if (selectedSessionId) {
      setLoading(true);
      mepApi(`/api/mep/review/session/${selectedSessionId}`)
        .then(res => setSessionData(res))
        .catch(err => console.error(err))
        .finally(() => setLoading(false));
    }
  }, [selectedSessionId]);

  const sess = sessionData?.session;
  const items = sessionData?.items || [];

  return (
    <div className="tab-page-container">
      <div className="page-header">
        <div>
          <h2>📈 Báo cáo & Xuất file Excel</h2>
          <p>Xuất file BOQ đã tô màu cảnh báo kèm ghi chú và Báo cáo chênh lệch đơn giá 16 cột chi tiết.</p>
        </div>
      </div>

      {/* Select Session */}
      <div className="report-session-selector">
        <label>
          Chọn phiên soát xét để xem và xuất báo cáo:
          <select value={selectedSessionId} onChange={e => setSelectedSessionId(e.target.value)}>
            {sessions.map(s => (
              <option key={s.id} value={s.id}>
                {s.original_filename || s.workbook_id} ({s.sheet_name}) — {s.created_at?.slice(0, 16).replace('T', ' ')}
              </option>
            ))}
          </select>
        </label>
      </div>

      {loading ? (
        <div className="table-loading">Đang tải báo cáo…</div>
      ) : sess ? (
        <div className="report-content-body">
          {/* Two Big Download Action Cards */}
          <div className="download-cards-grid">
            <div className="download-action-card">
              <div className="card-badge">FILE 1</div>
              <h3>📥 File BOQ sau soát xét (.xlsx)</h3>
              <p>
                Giữ nguyên 100% cấu trúc, công thức, định dạng của file Excel gốc. Các ô đơn giá được tô màu cảnh báo
                (🔴 Đỏ, 🟡 Vàng, 🟢 Xanh) kèm Excel Comment ghi rõ căn cứ và giá thư viện tham chiếu.
              </p>
              <a
                href={`/api/mep/workbooks/${sess.workbook_id}/download-reviewed?session_id=${sess.id}`}
                download
                className="primary-btn download-btn"
              >
                Tải file BOQ sau soát xét (.xlsx)
              </a>
            </div>

            <div className="download-action-card">
              <div className="card-badge">FILE 2</div>
              <h3>📊 Báo cáo chênh lệch đơn giá (.xlsx)</h3>
              <p>
                Bảng phân tích 16 cột chuẩn: STT, Sheet, Dòng Excel, Mã VT, Mô tả, ĐVT, Khối lượng, Đơn giá BOQ,
                Đơn giá Thư viện, Chênh lệch tiền (đ), Tỷ lệ (%), Nguồn giá, Kết quả kiểm tra, Trạng thái xử lý.
              </p>
              <a
                href={`/api/mep/reports/variance/${sess.id}`}
                download
                className="secondary-btn download-btn"
              >
                Tải Báo cáo chênh lệch (.xlsx)
              </a>
            </div>
          </div>

          {/* Statistical Breakdown */}
          <div className="report-stats-grid">
            <div className="report-stat-card">
              <h4>Tổng số đầu mục đã kiểm tra</h4>
              <div className="stat-big-num">{sess.total_items}</div>
              <small>Được lọc từ file BOQ (đã loại bỏ dòng tiêu đề và subtotal)</small>
            </div>

            <div className="report-stat-card">
              <h4>Tỷ lệ đạt chuẩn</h4>
              <div className="stat-big-num text-green">
                {sess.total_items > 0 ? Math.round((sess.passed_count / sess.total_items) * 100) : 0}%
              </div>
              <small>{sess.passed_count} / {sess.total_items} đầu mục nằm trong ngưỡng ±{sess.threshold_percent}%</small>
            </div>

            <div className="report-stat-card">
              <h4>Tổng chênh lệch thành tiền</h4>
              <div className={`stat-big-num ${sess.total_variance > 0 ? 'text-red' : 'text-green'}`}>
                {sess.total_variance ? `${sess.total_variance.toLocaleString('vi-VN')} đ` : '0 đ'}
              </div>
              <small>Tổng cộng chênh lệch giữa giá BOQ và giá thư viện × Khối lượng</small>
            </div>
          </div>

          {/* Detailed preview list */}
          <div className="data-table-container" style={{ marginTop: '24px' }}>
            <h4 style={{ padding: '16px 16px 8px 16px', margin: 0 }}>Xem trước danh sách chênh lệch:</h4>
            <table className="custom-table">
              <thead>
                <tr>
                  <th style={{ width: '40px' }}>#</th>
                  <th>Mô tả vật tư BOQ</th>
                  <th style={{ width: '60px' }}>ĐVT</th>
                  <th style={{ width: '80px', textAlign: 'right' }}>Khối lượng</th>
                  <th style={{ width: '120px', textAlign: 'right' }}>Giá BOQ (đ)</th>
                  <th style={{ width: '120px', textAlign: 'right' }}>Giá Thư viện (đ)</th>
                  <th style={{ width: '110px', textAlign: 'right' }}>Chênh lệch (đ)</th>
                  <th style={{ width: '80px', textAlign: 'right' }}>Tỷ lệ</th>
                  <th style={{ width: '150px' }}>Kết quả</th>
                  <th style={{ width: '150px' }}>Trạng thái xử lý</th>
                </tr>
              </thead>
              <tbody>
                {items.slice(0, 100).map((it: any, idx: number) => (
                  <tr key={it.id}>
                    <td className="text-center">{idx + 1}</td>
                    <td><strong>{it.boq_description}</strong></td>
                    <td className="text-center">{it.boq_unit || '—'}</td>
                    <td className="text-right">{it.boq_quantity ? it.boq_quantity.toLocaleString('vi-VN') : '—'}</td>
                    <td className="text-right text-bold">{it.boq_price.toLocaleString('vi-VN')} đ</td>
                    <td className="text-right text-bold">
                      {it.lib_price ? `${it.lib_price.toLocaleString('vi-VN')} đ` : '—'}
                    </td>
                    <td className={`text-right text-bold ${it.difference > 0 ? 'text-red' : it.difference < 0 ? 'text-green' : ''}`}>
                      {it.difference ? `${it.difference > 0 ? '+' : ''}${it.difference.toLocaleString('vi-VN')} đ` : '0 đ'}
                    </td>
                    <td className={`text-right text-bold ${it.diff_percent > 0 ? 'text-red' : it.diff_percent < 0 ? 'text-green' : ''}`}>
                      {it.diff_percent !== null ? `${it.diff_percent > 0 ? '+' : ''}${it.diff_percent}%` : '—'}
                    </td>
                    <td>
                      <span className={`status-pill badge-${it.warning_color}`}>
                        {it.status === 'higher' ? 'Giá cao hơn' : it.status === 'lower' ? 'Giá thấp hơn' : it.status === 'passed' ? 'Đạt' : 'Chưa có giá'}
                      </span>
                    </td>
                    <td>{it.user_action !== 'none' ? it.user_action : 'Chưa xử lý'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="table-empty">Chưa có phiên soát xét nào được chọn. Vui lòng chạy soát giá trước.</div>
      )}
    </div>
  );
}

