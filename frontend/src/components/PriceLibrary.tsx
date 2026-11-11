import React, { useState, useEffect } from 'react';
import { mepApi } from '../mepApi';

interface LibraryItem {
  id: number;
  system: string;
  code: string;
  description: string;
  specification: string;
  brand: string;
  origin: string;
  model: string;
  unit: string;
  material_price: number | null;
  labor_price: number | null;
  total_price: number;
  source: string;
  vendor: string;
  project: string;
  date: string;
  version: number;
}

export function PriceLibrary() {
  const [items, setItems] = useState<LibraryItem[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedSystem, setSelectedSystem] = useState('Tất cả');

  // Add Item Modal
  const [showAddModal, setShowAddModal] = useState(false);
  const [newItem, setNewItem] = useState({
    system: 'Điện',
    code: '',
    description: '',
    specification: '',
    brand: '',
    model: '',
    unit: 'Cái',
    total_price: '',
    source: 'Báo giá nhà cung cấp',
    vendor: '',
    project: ''
  });

  // Upload Library File
  const [uploading, setUploading] = useState(false);
  const [uploadPreview, setUploadPreview] = useState<any>(null);

  const systems = ['Tất cả', 'Điện', 'Cấp thoát nước', 'HVAC', 'PCCC', 'Khác'];

  const fetchItems = async () => {
    setLoading(true);
    try {
      const data = await mepApi(
        `/api/mep/libraries?system=${encodeURIComponent(selectedSystem)}&query=${encodeURIComponent(searchQuery)}&limit=150`
      );
      setItems(data.items || []);
      setTotal(data.total || 0);
    } catch (err: any) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchItems();
  }, [selectedSystem, searchQuery]);

  const handleAddItem = async (e: React.FormEvent) => {
    e.preventDefault();
    const priceNum = parseFloat(newItem.total_price.replace(/,/g, ''));
    if (!newItem.description.trim() || isNaN(priceNum) || priceNum <= 0) {
      alert('Vui lòng nhập tên/mô tả và đơn giá hợp lệ (> 0)');
      return;
    }

    try {
      await mepApi('/api/mep/libraries/item', {
        method: 'POST',
        body: JSON.stringify({
          ...newItem,
          total_price: priceNum
        })
      });
      setShowAddModal(false);
      setNewItem({
        system: 'Điện',
        code: '',
        description: '',
        specification: '',
        brand: '',
        model: '',
        unit: 'Cái',
        total_price: '',
        source: 'Báo giá nhà cung cấp',
        vendor: '',
        project: ''
      });
      fetchItems();
      alert('Đã thêm vật tư vào thư viện thành công!');
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    }
  };

  const handleQuickEditPrice = async (item: LibraryItem) => {
    const input = prompt(`Cập nhật đơn giá cho "${item.description}" (hiện tại: ${item.total_price.toLocaleString('vi-VN')} đ):`, String(item.total_price));
    if (!input) return;
    const newPrice = parseFloat(input.replace(/,/g, ''));
    if (isNaN(newPrice) || newPrice <= 0) {
      alert('Đơn giá không hợp lệ');
      return;
    }
    const reason = prompt('Nhập lý do điều chỉnh đơn giá:', 'Cập nhật theo bảng giá mới');

    try {
      await mepApi(`/api/mep/libraries/item/${item.id}`, {
        method: 'PUT',
        body: JSON.stringify({
          ...item,
          total_price: newPrice,
          reason: reason || 'Cập nhật đơn giá'
        })
      });
      fetchItems();
      alert('Đã cập nhật đơn giá và lưu lịch sử thay đổi!');
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    }
  };

  const handleDeleteItem = async (item: LibraryItem) => {
    if (!confirm(`Bạn có chắc muốn xóa "${item.description}" khỏi thư viện?`)) return;
    try {
      await mepApi(`/api/mep/libraries/item/${item.id}`, { method: 'DELETE' });
      fetchItems();
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    }
  };

  const handleUploadFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    const form = new FormData();
    form.append('file', file);
    try {
      const data = await mepApi('/api/mep/libraries/upload', {
        method: 'POST',
        body: form
      });
      setUploadPreview(data);
    } catch (err: any) {
      alert('Lỗi tải file: ' + err.message);
    } finally {
      setUploading(false);
      e.target.value = '';
    }
  };

  return (
    <div className="tab-page-container">
      <div className="page-header">
        <div>
          <h2>📚 Thư viện đơn giá MEP</h2>
          <p>Quản lý nguồn giá vật tư, thiết bị MEP theo từng phân hệ. Lưu lịch sử và phiên bản giá.</p>
        </div>
        <div className="header-actions">
          <label className="file-upload-btn">
            {uploading ? 'Đang nạp file…' : '📂 Nạp file Excel thư viện'}
            <input type="file" accept=".xlsx,.xlsm" onChange={handleUploadFile} disabled={uploading} />
          </label>
          <button className="primary-btn" onClick={() => setShowAddModal(true)}>
            ＋ Thêm vật tư mới
          </button>
        </div>
      </div>

      {/* System Filter Pills */}
      <div className="filter-pill-bar">
        {systems.map(sys => (
          <button
            key={sys}
            className={`pill-btn ${selectedSystem === sys ? 'active' : ''}`}
            onClick={() => setSelectedSystem(sys)}
          >
            {sys}
          </button>
        ))}
      </div>

      {/* Search and Stats bar */}
      <div className="table-search-bar">
        <input
          type="text"
          placeholder="Tìm theo tên vật tư, mã hiệu, thông số, nhãn hiệu hoặc model..."
          value={searchQuery}
          onChange={e => setSearchQuery(e.target.value)}
          className="search-input-large"
        />
        <div className="stats-badge">
          Tìm thấy: <strong>{total.toLocaleString('vi-VN')}</strong> đầu mục
        </div>
      </div>

      {/* Library Table */}
      <div className="data-table-container">
        {loading ? (
          <div className="table-loading">Đang tải dữ liệu thư viện…</div>
        ) : items.length === 0 ? (
          <div className="table-empty">
            <p>Không có vật tư nào phù hợp với điều kiện tìm kiếm.</p>
          </div>
        ) : (
          <table className="custom-table">
            <thead>
              <tr>
                <th style={{ width: '40px' }}>#</th>
                <th style={{ width: '90px' }}>Hệ MEP</th>
                <th style={{ width: '90px' }}>Mã VT</th>
                <th>Mô tả / Tên vật tư</th>
                <th>Thông số kỹ thuật</th>
                <th style={{ width: '100px' }}>Nhãn hiệu</th>
                <th style={{ width: '60px' }}>ĐVT</th>
                <th style={{ width: '130px', textAlign: 'right' }}>Đơn giá (đ)</th>
                <th style={{ width: '160px' }}>Nguồn giá</th>
                <th style={{ width: '90px' }}>Ngày</th>
                <th style={{ width: '110px', textAlign: 'center' }}>Thao tác</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, idx) => (
                <tr key={item.id}>
                  <td className="text-center">{idx + 1}</td>
                  <td>
                    <span className={`system-badge sys-${item.system}`}>{item.system}</span>
                  </td>
                  <td><code>{item.code || '—'}</code></td>
                  <td><strong>{item.description}</strong></td>
                  <td><small>{item.specification || '—'}</small></td>
                  <td>{item.brand || '—'}</td>
                  <td className="text-center">{item.unit || '—'}</td>
                  <td className="text-right text-bold">
                    {item.total_price.toLocaleString('vi-VN')} đ
                  </td>
                  <td><small title={item.source}>{item.source || '—'}</small></td>
                  <td><small>{item.date || '—'}</small></td>
                  <td className="text-center">
                    <button
                      className="inline-btn edit-btn"
                      title="Sửa đơn giá"
                      onClick={() => handleQuickEditPrice(item)}
                    >
                      ✏️ Sửa
                    </button>
                    <button
                      className="inline-btn del-btn"
                      title="Xóa"
                      onClick={() => handleDeleteItem(item)}
                    >
                      🗑
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Add Item Modal */}
      {showAddModal && (
        <div className="modal-backdrop">
          <div className="modal-card">
            <div className="modal-header">
              <h3>Thêm vật tư vào thư viện đơn giá MEP</h3>
              <button className="close-btn" onClick={() => setShowAddModal(false)}>×</button>
            </div>
            <form onSubmit={handleAddItem} className="modal-form">
              <div className="form-grid">
                <label>
                  Hệ MEP *
                  <select
                    value={newItem.system}
                    onChange={e => setNewItem({ ...newItem, system: e.target.value })}
                  >
                    <option value="Điện">Điện</option>
                    <option value="Cấp thoát nước">Cấp thoát nước</option>
                    <option value="HVAC">Điều hòa thông gió (HVAC)</option>
                    <option value="PCCC">Phòng cháy chữa cháy (PCCC)</option>
                    <option value="Khác">Khác</option>
                  </select>
                </label>

                <label>
                  Mã vật tư
                  <input
                    type="text"
                    placeholder="Ví dụ: CV-4, MCCB-1600..."
                    value={newItem.code}
                    onChange={e => setNewItem({ ...newItem, code: e.target.value })}
                  />
                </label>
              </div>

              <label>
                Mô tả / Tên vật tư *
                <input
                  type="text"
                  required
                  placeholder="Ví dụ: Cáp điện CV 1x4mm2..."
                  value={newItem.description}
                  onChange={e => setNewItem({ ...newItem, description: e.target.value })}
                />
              </label>

              <label>
                Thông số kỹ thuật
                <input
                  type="text"
                  placeholder="Ví dụ: 1x4mm2 - 0.6/1kV Cu/PVC..."
                  value={newItem.specification}
                  onChange={e => setNewItem({ ...newItem, specification: e.target.value })}
                />
              </label>

              <div className="form-grid">
                <label>
                  Nhãn hiệu / Hãng
                  <input
                    type="text"
                    placeholder="Ví dụ: Cadivi, Schneider..."
                    value={newItem.brand}
                    onChange={e => setNewItem({ ...newItem, brand: e.target.value })}
                  />
                </label>

                <label>
                  Model
                  <input
                    type="text"
                    placeholder="Ví dụ: NF1600-SEW..."
                    value={newItem.model}
                    onChange={e => setNewItem({ ...newItem, model: e.target.value })}
                  />
                </label>
              </div>

              <div className="form-grid">
                <label>
                  Đơn vị tính *
                  <input
                    type="text"
                    required
                    placeholder="Ví dụ: m, Cái, Bộ..."
                    value={newItem.unit}
                    onChange={e => setNewItem({ ...newItem, unit: e.target.value })}
                  />
                </label>

                <label>
                  Đơn giá tổng hợp (VNĐ) *
                  <input
                    type="text"
                    required
                    placeholder="Ví dụ: 14200, 41697600..."
                    value={newItem.total_price}
                    onChange={e => setNewItem({ ...newItem, total_price: e.target.value })}
                  />
                </label>
              </div>

              <div className="form-grid">
                <label>
                  Nguồn giá / Chứng chỉ
                  <input
                    type="text"
                    value={newItem.source}
                    onChange={e => setNewItem({ ...newItem, source: e.target.value })}
                  />
                </label>

                <label>
                  Nhà cung cấp
                  <input
                    type="text"
                    placeholder="Tên nhà cung cấp..."
                    value={newItem.vendor}
                    onChange={e => setNewItem({ ...newItem, vendor: e.target.value })}
                  />
                </label>
              </div>

              <div className="modal-actions">
                <button type="button" className="tool-btn" onClick={() => setShowAddModal(false)}>
                  Hủy bỏ
                </button>
                <button type="submit" className="primary-btn">
                  Lưu vào thư viện
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

