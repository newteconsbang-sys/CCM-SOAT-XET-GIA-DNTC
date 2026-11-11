import React, { useState, useEffect } from 'react';
import { mepApi } from '../mepApi';

interface ReviewHistoryProps {
  onSelectSession?: (sessionId: string) => void;
}

export function ReviewHistory({ onSelectSession }: ReviewHistoryProps) {
  const [activeTab, setActiveTab] = useState<'sessions' | 'memory'>('sessions');
  const [sessions, setSessions] = useState<any[]>([]);
  const [learningMemory, setLearningMemory] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);

  const fetchData = async () => {
    setLoading(true);
    try {
      const sessData = await mepApi('/api/mep/history/sessions');
      setSessions(sessData || []);
      const memData = await mepApi('/api/mep/review/learning-memory');
      setLearningMemory(memData || []);
    } catch (err: any) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const handleRevokeMemory = async (id: number) => {
    if (!confirm('Bạn có chắc muốn hủy bản ghi nhớ tương đương này?')) return;
    try {
      await mepApi(`/api/mep/review/learning-memory/${id}`, { method: 'DELETE' });
      fetchData();
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    }
  };

  return (
    <div className="tab-page-container">
      <div className="page-header">
        <div>
          <h2>📜 Lịch sử soát xét & Bộ nhớ ghi nhớ</h2>
          <p>Xem lại lịch sử các đợt soát giá BOQ và quản lý các cặp vật tư tương đương đã được xác nhận.</p>
        </div>
      </div>

      {/* Sub Tabs */}
      <div className="filter-pill-bar">
        <button
          className={`pill-btn ${activeTab === 'sessions' ? 'active' : ''}`}
          onClick={() => setActiveTab('sessions')}
        >
          📁 Lịch sử các phiên soát giá ({sessions.length})
        </button>
        <button
          className={`pill-btn ${activeTab === 'memory' ? 'active' : ''}`}
          onClick={() => setActiveTab('memory')}
        >
          🧠 Bộ nhớ tương đương đã duyệt ({learningMemory.length})
        </button>
      </div>

      {loading ? (
        <div className="table-loading">Đang tải dữ liệu lịch sử…</div>
      ) : activeTab === 'sessions' ? (
        <div className="data-table-container">
          {sessions.length === 0 ? (
            <div className="table-empty">Chưa có phiên soát giá nào được thực hiện.</div>
          ) : (
            <table className="custom-table">
              <thead>
                <tr>
                  <th style={{ width: '40px' }}>#</th>
                  <th style={{ width: '160px' }}>Thời gian</th>
                  <th>File BOQ</th>
                  <th style={{ width: '120px' }}>Sheet</th>
                  <th style={{ width: '80px', textAlign: 'center' }}>Ngưỡng</th>
                  <th style={{ width: '80px', textAlign: 'center' }}>Tổng mục</th>
                  <th style={{ width: '70px', textAlign: 'center' }}>🟢 Đạt</th>
                  <th style={{ width: '70px', textAlign: 'center' }}>🔴 Cao</th>
                  <th style={{ width: '70px', textAlign: 'center' }}>🟡 Thấp</th>
                  <th style={{ width: '130px', textAlign: 'right' }}>Tổng chênh lệch</th>
                  <th style={{ width: '140px', textAlign: 'center' }}>Báo cáo</th>
                </tr>
              </thead>
              <tbody>
                {sessions.map((sess, idx) => (
                  <tr key={sess.id}>
                    <td className="text-center">{idx + 1}</td>
                    <td><small>{sess.created_at ? sess.created_at.slice(0, 19).replace('T', ' ') : '—'}</small></td>
                    <td><strong>{sess.original_filename || sess.workbook_id}</strong></td>
                    <td>{sess.sheet_name}</td>
                    <td className="text-center">±{sess.threshold_percent}%</td>
                    <td className="text-center text-bold">{sess.total_items}</td>
                    <td className="text-center text-green text-bold">{sess.passed_count}</td>
                    <td className="text-center text-red text-bold">{sess.higher_count}</td>
                    <td className="text-center text-yellow text-bold">{sess.lower_count}</td>
                    <td className={`text-right text-bold ${sess.total_variance > 0 ? 'text-red' : 'text-green'}`}>
                      {sess.total_variance ? `${sess.total_variance.toLocaleString('vi-VN')} đ` : '0 đ'}
                    </td>
                    <td className="text-center">
                      <a
                        href={`/api/mep/reports/variance/${sess.id}`}
                        download
                        className="inline-btn view-btn"
                        title="Tải file Excel báo cáo chênh lệch"
                      >
                        📥 Báo cáo Excel
                      </a>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      ) : (
        <div className="data-table-container">
          {learningMemory.length === 0 ? (
            <div className="table-empty">
              Chưa có cặp vật tư tương đương nào trong bộ nhớ ghi nhớ. Khi soát xét, chọn "Ghi nhớ cặp tương đương này" để hệ thống tự động nhận diện trong các lần sau.
            </div>
          ) : (
            <table className="custom-table">
              <thead>
                <tr>
                  <th style={{ width: '40px' }}>#</th>
                  <th>Mô tả vật tư BOQ</th>
                  <th style={{ width: '60px' }}>ĐVT</th>
                  <th>Vật tư Thư viện tương đương</th>
                  <th style={{ width: '130px', textAlign: 'right' }}>Giá xác nhận</th>
                  <th>Ghi chú căn cứ</th>
                  <th style={{ width: '130px' }}>Thời gian duyệt</th>
                  <th style={{ width: '80px', textAlign: 'center' }}>Thao tác</th>
                </tr>
              </thead>
              <tbody>
                {learningMemory.map((mem, idx) => (
                  <tr key={mem.id}>
                    <td className="text-center">{idx + 1}</td>
                    <td><strong>{mem.boq_description}</strong></td>
                    <td className="text-center">{mem.boq_unit || '—'}</td>
                    <td><strong>{mem.library_description}</strong></td>
                    <td className="text-right text-bold">
                      {mem.confirmed_price ? `${mem.confirmed_price.toLocaleString('vi-VN')} đ` : '—'}
                    </td>
                    <td><small>{mem.notes || '—'}</small></td>
                    <td><small>{mem.created_at ? mem.created_at.slice(0, 10) : '—'}</small></td>
                    <td className="text-center">
                      <button
                        className="inline-btn del-btn"
                        title="Hủy ghi nhớ tương đương này"
                        onClick={() => handleRevokeMemory(mem.id)}
                      >
                        Hủy
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </div>
  );
}

