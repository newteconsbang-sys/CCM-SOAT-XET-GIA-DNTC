import React, { useState, useEffect } from 'react';
import { mepApi } from '../mepApi';

interface PriceReviewProps {
  onOpenInExcel: (workbookId: string, sheetName: string, row: number, col: number) => void;
  activeWorkbookId: string | null;
  onWorkbookSelected: (wbId: string) => void;
}

export function PriceReview({
  onOpenInExcel,
  activeWorkbookId,
  onWorkbookSelected
}: PriceReviewProps) {
  const [workbooks, setWorkbooks] = useState<any[]>([]);
  const [selectedWbId, setSelectedWbId] = useState<string>(activeWorkbookId || '');
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [headerRow, setHeaderRow] = useState<number>(2);

  // 9 Mapping Columns
  const [mapping, setMapping] = useState({
    stt: null as number | null,
    code: null as number | null,
    description: null as number | null,
    specification: null as number | null,
    brand: null as number | null,
    model: null as number | null,
    unit: null as number | null,
    quantity: null as number | null,
    price: null as number | null,
    amount: null as number | null,
  });

  const [headersList, setHeadersList] = useState<Record<number, string>>({});
  const [thresholdPercent, setThresholdPercent] = useState<number>(10.0);

  // Review Execution State
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null);
  const [reviewSummary, setReviewSummary] = useState<any>(null);
  const [reviewItems, setReviewItems] = useState<any[]>([]);

  // Filter & Search
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [uploading, setUploading] = useState(false);

  // Load uploaded workbooks
  const loadWorkbooks = async () => {
    try {
      const data = await mepApi('/api/mep/workbooks');
      setWorkbooks(data || []);
      if (!selectedWbId && data.length > 0) {
        setSelectedWbId(data[0].id);
        onWorkbookSelected(data[0].id);
      }
    } catch (err: any) {
      console.error(err);
    }
  };

  useEffect(() => {
    loadWorkbooks();
  }, []);

  // Update sheets when workbook changes
  useEffect(() => {
    if (selectedWbId) {
      const wb = workbooks.find(w => w.id === selectedWbId);
      if (wb && wb.sheets) {
        const sNames = wb.sheets.map((s: any) => s.name);
        setSheets(sNames);
        if (sNames.length > 0 && !sNames.includes(selectedSheet)) {
          setSelectedSheet(sNames[0]);
        }
      }
    }
  }, [selectedWbId, workbooks]);

  // Propose mapping when workbook or sheet changes
  const autoDetectMapping = async () => {
    if (!selectedWbId || !selectedSheet) return;
    try {
      const res = await mepApi('/api/mep/review/suggest-mapping', {
        method: 'POST',
        body: JSON.stringify({
          workbook_id: selectedWbId,
          sheet_name: selectedSheet,
          header_row: headerRow
        })
      });
      setHeaderRow(res.header_row || 2);
      setMapping(res.mapping || {});
      setHeadersList(res.headers || {});
    } catch (err: any) {
      console.error(err);
    }
  };

  useEffect(() => {
    if (selectedWbId && selectedSheet) {
      autoDetectMapping();
    }
  }, [selectedWbId, selectedSheet]);

  // Upload new BOQ
  const handleUploadBoq = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    const form = new FormData();
    form.append('file', file);
    try {
      const res = await mepApi('/api/mep/workbooks/upload', {
        method: 'POST',
        body: form
      });
      await loadWorkbooks();
      setSelectedWbId(res.id);
      onWorkbookSelected(res.id);
      if (res.sheets && res.sheets.length > 0) {
        setSelectedSheet(res.sheets[0].name);
      }
      alert(`Đã nạp file ${res.filename} thành công!`);
    } catch (err: any) {
      alert('Lỗi nạp file: ' + err.message);
    } finally {
      setUploading(false);
      e.target.value = '';
    }
  };

  // Run price review
  const handleRunReview = async () => {
    if (!mapping.description || !mapping.price) {
      alert('Vui lòng chọn cột "Mô tả / Tên vật tư" và cột "Đơn giá" trước khi chạy soát giá.');
      return;
    }

    setRunning(true);
    setProgress(20);

    try {
      setProgress(50);
      const res = await mepApi('/api/mep/review/run', {
        method: 'POST',
        body: JSON.stringify({
          workbook_id: selectedWbId,
          sheet_name: selectedSheet,
          header_row: headerRow,
          mapping,
          threshold_percent: thresholdPercent
        })
      });

      setProgress(85);
      setCurrentSessionId(res.session_id);
      setReviewSummary(res.summary);

      // Fetch items
      const details = await mepApi(`/api/mep/review/session/${res.session_id}`);
      setReviewItems(details.items || []);
      setProgress(100);
    } catch (err: any) {
      alert('Lỗi chạy soát giá: ' + err.message);
    } finally {
      setRunning(false);
    }
  };

  // 1-Click action from table
  const handleQuickAction = async (item: any, action: 'applied_lib' | 'kept_boq') => {
    try {
      await mepApi(`/api/mep/review/item/${item.id}/action`, {
        method: 'POST',
        body: JSON.stringify({
          action,
          user_note: action === 'applied_lib' ? 'Áp dụng giá thư viện' : 'Giữ nguyên giá BOQ',
          remember_memory: true
        })
      });
      // update state
      setReviewItems(prev =>
        prev.map(it => (it.id === item.id ? { ...it, user_action: action, warning_color: 'green' } : it))
      );
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    }
  };

  // Filtered rows
  const filteredItems = reviewItems.filter(item => {
    if (statusFilter !== 'all') {
      if (statusFilter === 'higher' && item.status !== 'higher') return false;
      if (statusFilter === 'lower' && item.status !== 'lower') return false;
      if (statusFilter === 'passed' && item.status !== 'passed') return false;
      if (statusFilter === 'not_found' && item.status !== 'not_found') return false;
      if (statusFilter === 'spec_mismatch' && item.status !== 'spec_mismatch') return false;
    }
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      const text = `${item.boq_description} ${item.boq_code || ''} ${item.lib_description || ''}`.toLowerCase();
      return text.includes(q);
    }
    return true;
  });

  return (
    <div className="tab-page-container">
      <div className="page-header">
        <div>
          <h2>⚡ Soát xét đơn giá MEP</h2>
          <p>Đối chiếu tự động giữa file dự toán BOQ và Thư viện đơn giá theo các quy tắc nghiệp vụ MEP chuẩn hóa.</p>
        </div>
        <div className="header-actions">
          <label className="file-upload-btn">
            {uploading ? 'Đang nạp…' : '📁 Tải lên file BOQ (.xlsx, .xlsm)'}
            <input type="file" accept=".xlsx,.xlsm" onChange={handleUploadBoq} disabled={uploading} />
          </label>
        </div>
      </div>

      {/* Configuration Card */}
      <div className="config-card">
        <div className="config-header">
          <h3>1. Chọn file BOQ và cấu hình Mapping cột</h3>
          <button className="tool-btn" onClick={autoDetectMapping}>
            🔄 Tự động nhận diện lại cột
          </button>
        </div>

        <div className="config-grid-3">
          <label>
            File BOQ đã nạp
            <select
              value={selectedWbId}
              onChange={e => {
                setSelectedWbId(e.target.value);
                onWorkbookSelected(e.target.value);
              }}
            >
              {workbooks.map(wb => (
                <option key={wb.id} value={wb.id}>
                  {wb.original_filename} ({wb.file_type})
                </option>
              ))}
            </select>
          </label>

          <label>
            Sheet cần soát giá
            <select value={selectedSheet} onChange={e => setSelectedSheet(e.target.value)}>
              {sheets.map(s => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </label>

          <label>
            Dòng tiêu đề (Header Row)
            <input
              type="number"
              min={1}
              value={headerRow}
              onChange={e => setHeaderRow(parseInt(e.target.value) || 1)}
            />
          </label>
        </div>

        {/* 9 Column Mapping Selectors */}
        <h4 style={{ marginTop: '16px', marginBottom: '8px' }}>Ánh xạ các cột trong file Excel BOQ:</h4>
        <div className="mapping-columns-grid">
          {[
            ['code', 'Mã vật tư'],
            ['description', 'Mô tả / Tên vật tư *'],
            ['specification', 'Thông số kỹ thuật'],
            ['brand', 'Nhãn hiệu'],
            ['model', 'Model'],
            ['unit', 'Đơn vị tính'],
            ['quantity', 'Khối lượng'],
            ['price', 'Đơn giá *'],
            ['amount', 'Thành tiền'],
          ].map(([key, label]) => (
            <div key={key} className="mapping-column-box">
              <label>{label}</label>
              <select
                value={mapping[key as keyof typeof mapping] ?? ''}
                onChange={e =>
                  setMapping({
                    ...mapping,
                    [key]: e.target.value === '' ? null : parseInt(e.target.value)
                  })
                }
              >
                <option value="">-- Chưa chọn --</option>
                {Object.entries(headersList).map(([colIdx, title]) => (
                  <option key={colIdx} value={colIdx}>
                    Cột {colIdx}: {title}
                  </option>
                ))}
              </select>
            </div>
          ))}
        </div>

        {/* Threshold selection and Run CTA */}
        <div className="run-action-bar">
          <div className="threshold-selector">
            <span>Ngưỡng cảnh báo chênh lệch:</span>
            {[5, 10, 15, 20].map(val => (
              <button
                key={val}
                className={`thresh-btn ${thresholdPercent === val ? 'active' : ''}`}
                onClick={() => setThresholdPercent(val)}
              >
                ±{val}%
              </button>
            ))}
            <div className="custom-thresh">
              <input
                type="number"
                min={1}
                max={100}
                value={thresholdPercent}
                onChange={e => setThresholdPercent(parseFloat(e.target.value) || 10)}
              />
              <span>%</span>
            </div>
          </div>

          <button
            className="primary-btn large-cta"
            disabled={running || !selectedWbId}
            onClick={handleRunReview}
          >
            {running ? `Đang soát giá… (${progress}%)` : '🚀 CHẠY SOÁT GIÁ'}
          </button>
        </div>

        {running && (
          <div className="progress-bar-container">
            <div className="progress-fill" style={{ width: `${progress}%` }} />
          </div>
        )}
      </div>

      {/* Review Results Section */}
      {reviewSummary && (
        <div className="review-results-section">
          {/* KPI Dashboard */}
          <div className="summary-kpi-cards">
            <div className="kpi-card total-card">
              <span className="kpi-label">Tổng đầu mục</span>
              <strong className="kpi-value">{reviewSummary.total}</strong>
            </div>

            <div className="kpi-card green-card">
              <span className="kpi-label">🟢 Đạt (trong ngưỡng)</span>
              <strong className="kpi-value">{reviewSummary.passed}</strong>
            </div>

            <div className="kpi-card red-card">
              <span className="kpi-label">🔴 Giá cao hơn</span>
              <strong className="kpi-value">{reviewSummary.higher}</strong>
            </div>

            <div className="kpi-card yellow-card">
              <span className="kpi-label">🟡 Giá thấp hơn</span>
              <strong className="kpi-value">{reviewSummary.lower}</strong>
            </div>

            <div className="kpi-card gray-card">
              <span className="kpi-label">⚪ Chưa có giá</span>
              <strong className="kpi-value">{reviewSummary.not_found}</strong>
            </div>

            <div className="kpi-card alert-card">
              <span className="kpi-label">⚠️ Lệch thông số</span>
              <strong className="kpi-value">{reviewSummary.spec_mismatch}</strong>
            </div>

            <div className="kpi-card diff-card">
              <span className="kpi-label">Tổng chênh lệch dự kiến</span>
              <strong className={`kpi-value ${reviewSummary.total_variance > 0 ? 'text-red' : 'text-green'}`}>
                {reviewSummary.total_variance.toLocaleString('vi-VN')} đ
              </strong>
            </div>
          </div>

          {/* Table Filter Tabs */}
          <div className="results-filter-toolbar">
            <div className="filter-tabs">
              {[
                ['all', `Tất cả (${reviewItems.length})`],
                ['higher', `🔴 Cao hơn (${reviewSummary.higher})`],
                ['lower', `🟡 Thấp hơn (${reviewSummary.lower})`],
                ['passed', `🟢 Đạt (${reviewSummary.passed})`],
                ['spec_mismatch', `⚠️ Lệch thông số (${reviewSummary.spec_mismatch})`],
                ['not_found', `⚪ Chưa có giá (${reviewSummary.not_found})`],
              ].map(([key, label]) => (
                <button
                  key={key}
                  className={`tab-filter-btn ${statusFilter === key ? 'active' : ''}`}
                  onClick={() => setStatusFilter(key)}
                >
                  {label}
                </button>
              ))}
            </div>

            <div className="search-box">
              <input
                type="text"
                placeholder="Lọc trong kết quả..."
                value={searchQuery}
                onChange={e => setSearchQuery(e.target.value)}
              />
            </div>
          </div>

          {/* Detailed Results Table */}
          <div className="data-table-container">
            <table className="custom-table">
              <thead>
                <tr>
                  <th style={{ width: '40px' }}>#</th>
                  <th style={{ width: '70px' }}>Dòng</th>
                  <th>Mô tả vật tư BOQ</th>
                  <th style={{ width: '60px' }}>ĐVT</th>
                  <th style={{ width: '80px', textAlign: 'right' }}>Khối lượng</th>
                  <th style={{ width: '120px', textAlign: 'right' }}>Đơn giá BOQ</th>
                  <th style={{ width: '120px', textAlign: 'right' }}>Giá Thư viện</th>
                  <th style={{ width: '110px', textAlign: 'right' }}>Chênh lệch</th>
                  <th style={{ width: '90px', textAlign: 'right' }}>Tỷ lệ (%)</th>
                  <th style={{ width: '150px' }}>Kết quả kiểm tra</th>
                  <th style={{ width: '160px', textAlign: 'center' }}>Thao tác</th>
                </tr>
              </thead>
              <tbody>
                {filteredItems.map((item, idx) => {
                  let statusBadge = '';
                  let statusText = '';
                  if (item.status === 'higher') {
                    statusBadge = 'badge-red';
                    statusText = '🔴 Giá cao hơn';
                  } else if (item.status === 'lower') {
                    statusBadge = 'badge-yellow';
                    statusText = '🟡 Giá thấp hơn';
                  } else if (item.status === 'passed') {
                    statusBadge = 'badge-green';
                    statusText = '🟢 Đạt';
                  } else if (item.status === 'spec_mismatch') {
                    statusBadge = 'badge-yellow';
                    statusText = '⚠️ Lệch thông số';
                  } else {
                    statusBadge = 'badge-gray';
                    statusText = '⚪ Chưa có giá';
                  }

                  return (
                    <tr key={item.id} className={item.warning_color ? `row-${item.warning_color}` : ''}>
                      <td className="text-center">{idx + 1}</td>
                      <td className="text-center"><code>Dòng {item.excel_row}</code></td>
                      <td>
                        <strong>{item.boq_description}</strong>
                        {item.lib_description && item.lib_description !== item.boq_description && (
                          <div className="ref-subtext">
                            ↳ Khớp với: <em>{item.lib_description}</em>
                          </div>
                        )}
                        {item.user_note && <div className="note-subtext">Ghi chú: {item.user_note}</div>}
                      </td>
                      <td className="text-center">{item.boq_unit || '—'}</td>
                      <td className="text-right">{item.boq_quantity ? item.boq_quantity.toLocaleString('vi-VN') : '—'}</td>
                      <td className="text-right text-bold">
                        {item.boq_price.toLocaleString('vi-VN')} đ
                      </td>
                      <td className="text-right text-bold">
                        {item.lib_price ? `${item.lib_price.toLocaleString('vi-VN')} đ` : '—'}
                      </td>
                      <td className={`text-right text-bold ${item.difference > 0 ? 'text-red' : item.difference < 0 ? 'text-green' : ''}`}>
                        {item.difference !== 0 ? `${item.difference > 0 ? '+' : ''}${item.difference.toLocaleString('vi-VN')} đ` : '0 đ'}
                      </td>
                      <td className={`text-right text-bold ${item.diff_percent > 0 ? 'text-red' : item.diff_percent < 0 ? 'text-green' : ''}`}>
                        {item.diff_percent !== null ? `${item.diff_percent > 0 ? '+' : ''}${item.diff_percent}%` : '—'}
                      </td>
                      <td>
                        <span className={`status-pill ${statusBadge}`}>{statusText}</span>
                      </td>
                      <td className="text-center">
                        <div className="table-row-actions">
                          <button
                            className="inline-btn view-btn"
                            title="Mở ô này trên Excel Workspace để xem chi tiết và chỉnh sửa"
                            onClick={() => onOpenInExcel(selectedWbId, selectedSheet, item.excel_row, item.excel_col)}
                          >
                            📊 Mở Excel
                          </button>
                          {item.lib_price && item.status !== 'passed' && (
                            <button
                              className="inline-btn apply-btn"
                              title="Áp dụng giá thư viện"
                              onClick={() => handleQuickAction(item, 'applied_lib')}
                            >
                              ✓ Giá TV
                            </button>
                          )}
                          {item.status !== 'passed' && (
                            <button
                              className="inline-btn keep-btn"
                              title="Giữ nguyên giá BOQ"
                              onClick={() => handleQuickAction(item, 'kept_boq')}
                            >
                              ⚡ Giữ giá
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

