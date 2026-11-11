import React, { useState, useEffect } from 'react';
import { mepApi } from '../mepApi';

interface OverviewTabProps {
  onNavigateTab: (tab: string) => void;
  onSelectWorkbook: (wbId: string) => void;
}

export function OverviewTab({ onNavigateTab, onSelectWorkbook }: OverviewTabProps) {
  const [stats, setStats] = useState<any>(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    setLoading(true);
    mepApi('/api/mep/overview/stats')
      .then(res => setStats(res))
      .catch(err => console.error(err))
      .finally(() => setLoading(false));
  }, []);

  return (
    <div className="tab-page-container">
      {/* Welcome Hero Banner */}
      <div className="hero-banner">
        <div className="hero-content">
          <h1>CCM-AI PRICE REVIEW</h1>
          <h2>Hệ thống Soát xét Đơn giá Cơ Điện MEP trên Web Local</h2>
          <p>
            Tự động đối chiếu hồ sơ dự toán BOQ với thư viện đơn giá MEP bằng thuật toán chuẩn hóa kỹ thuật và RapidFuzz.
            Mở, chỉnh sửa trực tiếp trên Excel Web Local, tô màu cảnh báo chênh lệch và xuất báo cáo chi tiết.
          </p>
          <div className="hero-cta-buttons">
            <button className="primary-btn large" onClick={() => onNavigateTab('review')}>
              ⚡ Bắt đầu Soát xét đơn giá
            </button>
            <button className="secondary-btn large" onClick={() => onNavigateTab('workspace')}>
              📊 Mở Excel Workspace
            </button>
            <button className="outline-btn large" onClick={() => onNavigateTab('library')}>
              📚 Quản lý Thư viện giá
            </button>
          </div>
        </div>
      </div>

      {/* Overview Stat Cards */}
      <div className="overview-stats-grid">
        <div className="ov-card">
          <div className="ov-icon">📚</div>
          <div className="ov-info">
            <span className="ov-label">Thư viện đơn giá MEP</span>
            <strong className="ov-value">{stats?.library_total?.toLocaleString('vi-VN') || 0}</strong>
            <small>Đầu mục vật tư đã lưu trong SQLite</small>
          </div>
        </div>

        <div className="ov-card">
          <div className="ov-icon">📁</div>
          <div className="ov-info">
            <span className="ov-label">Hồ sơ BOQ đã nạp</span>
            <strong className="ov-value">{stats?.workbooks_total || 0}</strong>
            <small>File Excel (.xlsx, .xlsm)</small>
          </div>
        </div>

        <div className="ov-card">
          <div className="ov-icon">⚡</div>
          <div className="ov-info">
            <span className="ov-label">Số phiên soát giá</span>
            <strong className="ov-value">{stats?.sessions_total || 0}</strong>
            <small>Các đợt kiểm tra đã lưu lịch sử</small>
          </div>
        </div>

        <div className="ov-card">
          <div className="ov-icon">🧠</div>
          <div className="ov-info">
            <span className="ov-label">Bộ nhớ tương đương</span>
            <strong className="ov-value">{stats?.memory_total || 0}</strong>
            <small>Cặp vật tư người dùng đã duyệt</small>
          </div>
        </div>
      </div>

      {/* Distribution by System */}
      {stats?.library_by_system && Object.keys(stats.library_by_system).length > 0 && (
        <div className="system-breakdown-card">
          <h3>Phân bổ đơn giá thư viện theo hệ MEP</h3>
          <div className="system-pills-row">
            {Object.entries(stats.library_by_system).map(([sys, cnt]) => (
              <div key={sys} className="system-stat-pill">
                <span className="pill-name">{sys}</span>
                <strong className="pill-count">{Number(cnt).toLocaleString('vi-VN')} vật tư</strong>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Workflow Steps Guide */}
      <div className="workflow-guide-section">
        <h3>Quy trình soát xét đơn giá MEP chuẩn 5 bước</h3>
        <div className="workflow-steps-grid">
          <div className="wf-step-card">
            <div className="step-num">1</div>
            <h4>Nạp file BOQ</h4>
            <p>Upload file dự toán BOQ (.xlsx, .xlsm) cần soát xét. Bản gốc luôn được sao lưu an toàn.</p>
          </div>

          <div className="wf-step-card">
            <div className="step-num">2</div>
            <h4>Mapping cột</h4>
            <p>Hệ thống tự động nhận diện các cột Mã, Mô tả, ĐVT, Khối lượng, Đơn giá hoặc cho phép chọn thủ công.</p>
          </div>

          <div className="wf-step-card">
            <div className="step-num">3</div>
            <h4>Chạy soát giá tự động</h4>
            <p>Thuật toán Python + RapidFuzz đối chiếu mô tả, kiểm tra cấu hình kỹ thuật và tính tỷ lệ chênh lệch.</p>
          </div>

          <div className="wf-step-card">
            <div className="step-num">4</div>
            <h4>Xử lý trên Excel Web</h4>
            <p>Xem các ô tô màu đỏ/vàng trên Excel Web. Bấm vào ô để xem giá tham chiếu, áp dụng giá hoặc ghi chú.</p>
          </div>

          <div className="wf-step-card">
            <div className="step-num">5</div>
            <h4>Xuất báo cáo</h4>
            <p>Tải file BOQ đã tô màu kèm comment và Báo cáo chênh lệch đơn giá 16 cột phục vụ giải trình chủ đầu tư.</p>
          </div>
        </div>
      </div>
    </div>
  );
}

