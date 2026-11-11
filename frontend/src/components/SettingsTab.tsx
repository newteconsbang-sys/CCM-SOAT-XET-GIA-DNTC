import React, { useState } from 'react';

export function SettingsTab() {
  const [defaultThreshold, setDefaultThreshold] = useState(10);
  const [highlightColors, setHighlightColors] = useState({
    red: '#FEE2E2',
    yellow: '#FEF3C7',
    green: '#DCFCE7'
  });

  return (
    <div className="tab-page-container">
      <div className="page-header">
        <div>
          <h2>⚙️ Cài đặt & Quy tắc Soát xét MEP</h2>
          <p>Thiết lập ngưỡng cảnh báo mặc định, từ điển chuẩn hóa thuật ngữ MEP và nhóm đơn vị tính quy đổi.</p>
        </div>
      </div>

      <div className="settings-grid">
        {/* Threshold setting */}
        <div className="settings-card">
          <h3>1. Ngưỡng cảnh báo chênh lệch mặc định</h3>
          <p>Đơn giá BOQ vượt quá ngưỡng này so với Thư viện sẽ được đánh dấu 🔴 Đỏ (Vượt ngưỡng).</p>
          <div className="threshold-setting-box">
            <input
              type="range"
              min="1"
              max="50"
              value={defaultThreshold}
              onChange={e => setDefaultThreshold(parseInt(e.target.value))}
            />
            <span className="thresh-display">±{defaultThreshold}%</span>
          </div>
          <small>Các mốc thường dùng trong kiểm toán xây dựng: 5% (gói thầu lớn), 10% (tiêu chuẩn), 15% - 20% (vật tư đặc thù).</small>
        </div>

        {/* Color Palette settings */}
        <div className="settings-card">
          <h3>2. Bảng màu cảnh báo trên Excel</h3>
          <div className="color-palette-rows">
            <div className="color-row">
              <span className="color-preview" style={{ backgroundColor: highlightColors.red }} />
              <div>
                <strong>Màu đỏ (#FEE2E2):</strong>
                <p>Đơn giá BOQ cao hơn thư viện vượt ngưỡng cho phép.</p>
              </div>
            </div>

            <div className="color-row">
              <span className="color-preview" style={{ backgroundColor: highlightColors.yellow }} />
              <div>
                <strong>Màu vàng (#FEF3C7):</strong>
                <p>Đơn giá thấp hơn ngưỡng, chưa có giá tham chiếu hoặc không khớp thông số kỹ thuật.</p>
              </div>
            </div>

            <div className="color-row">
              <span className="color-preview" style={{ backgroundColor: highlightColors.green }} />
              <div>
                <strong>Màu xanh (#DCFCE7):</strong>
                <p>Đơn giá đạt trong phạm vi dung sai cho phép hoặc người dùng đã xác nhận.</p>
              </div>
            </div>
          </div>
        </div>

        {/* Normalization rules */}
        <div className="settings-card full-width">
          <h3>3. Bộ quy tắc chuẩn hóa thuật ngữ Cơ Điện MEP</h3>
          <p>Hệ thống tự động đồng nhất các cách viết tương đương trong dự toán MEP bằng Python regex & RapidFuzz:</p>

          <div className="rules-table-container">
            <table className="custom-table">
              <thead>
                <tr>
                  <th style={{ width: '150px' }}>Phân hệ MEP</th>
                  <th style={{ width: '250px' }}>Các cách viết tương đương</th>
                  <th>Thuật toán chuẩn hóa & Xử lý</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td><strong>Dây cáp điện</strong></td>
                  <td>Cáp CV 4.0mm2 = Dây điện CV 4mm² = Cáp điện CV 1x4mm2</td>
                  <td>
                    Quy chuẩn về <code>cap cv 1x4mm2</code>. Tách riêng tiết diện, số lõi, kiểm tra cấp điện áp
                    trước khi so sánh.
                  </td>
                </tr>
                <tr>
                  <td><strong>Thiết bị đóng cắt</strong></td>
                  <td>MCCB 3P-1600A -50KA = Cầu dao tự động MCCB 3P 1600A = Aptomat MCCB</td>
                  <td>
                    Chuẩn hóa thành <code>mccb 3p 1600a 50ka</code>. Tách số cực (3P), dòng định mức (1600A), dòng cắt ngắn mạch (50kA).
                  </td>
                </tr>
                <tr>
                  <td><strong>Đường ống MEP</strong></td>
                  <td>Ống PPR D25 = Ong PPR DN25 = Ống cấp nước PPR Ø25 PN16</td>
                  <td>
                    Đồng nhất đường kính (D25 = DN25 = Ø25), cấp áp lực (PN10, PN16, PN20).
                  </td>
                </tr>
                <tr>
                  <td><strong>Van công nghiệp</strong></td>
                  <td>Van cổng ty chìm = Van chặn = Gate valve; Van bướm tay gạt</td>
                  <td>
                    Nhận diện chuẩn loại van và kết nối (ren, bích, wafer).
                  </td>
                </tr>
                <tr>
                  <td><strong>Chiếu sáng</strong></td>
                  <td>Đèn LED âm trần = Đèn downlight = Đèn lon âm trần 18W</td>
                  <td>
                    So khớp công suất (W), nhiệt độ màu (3000K, 6500K), quang thông (lm).
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Unit Equivalence rules */}
        <div className="settings-card full-width">
          <h3>4. Bảng nhóm đơn vị tính quy đổi tương đương</h3>
          <p>Chỉ so sánh đơn giá khi hai vật tư thuộc cùng nhóm đơn vị tính hợp lệ:</p>
          <div className="unit-tags-grid">
            <div className="unit-group-box">
              <strong>Nhóm đơn vị: Cái / Bộ</strong>
              <span>Cái, Chiếc, Cây, Bộ, Quả, Hạt, Tấm</span>
            </div>
            <div className="unit-group-box">
              <strong>Nhóm chiều dài: Mét</strong>
              <span>m, mét, md, m dài</span>
            </div>
            <div className="unit-group-box">
              <strong>Nhóm diện tích: Mét vuông</strong>
              <span>m2, m², mét vuông</span>
            </div>
            <div className="unit-group-box">
              <strong>Nhóm thể tích: Mét khối</strong>
              <span>m3, m³, mét khối</span>
            </div>
            <div className="unit-group-box">
              <strong>Nhóm khối lượng: Kilogram</strong>
              <span>kg, kilogram, kí, tấn</span>
            </div>
            <div className="unit-group-box">
              <strong>Nhóm đóng gói: Cuộn / Hộp</strong>
              <span>cuộn, lô, hộp, thùng, gói</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

