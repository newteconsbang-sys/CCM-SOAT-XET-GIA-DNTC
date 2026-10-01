# CCM-AI – Soát xét giá

Excel task pane add-in cá nhân cho Microsoft 365 trên Windows. Giao diện React/TypeScript, Office.js đọc/ghi workbook; FastAPI/SQLite chạy HTTPS trên loopback; runtime Python được đóng gói riêng, không cần mở Codex khi sử dụng.

**Bắt đầu:** bấm `Cai-dat-CCM.cmd`, xác nhận chứng chỉ CCM-AI localhost trực tiếp trong Windows, rồi bấm shortcut Khởi động CCM-AI. Xem [hướng dẫn cài đặt](docs/CAI-DAT.md) và [hướng dẫn sử dụng](docs/HUONG-DAN.md).

**Trạng thái nghiệm thu:** xem [kết quả kiểm tra](docs/KET-QUA-KIEM-TRA.md). Đây là bản 0.1 để nghiệm thu cá nhân, chưa xác nhận triển khai sản xuất hoặc tích hợp Excel khi chưa có kết quả đọc/ghi Office.js thực tế.

## Các phần đã triển khai

- Tab/nút CCM-AI trong manifest; bảng điều khiển theo giao diện tham chiếu, bốn nghiệp vụ.
- Đọc hai vùng, giá trị/công thức/chuỗi hiển thị, danh sách sheet; ánh xạ linh hoạt, cột ĐVT/KL ngoài vùng, mẫu ánh xạ lưu SQLite.
- Ghép theo STT/mô tả/đơn vị với xác nhận khi thiếu hoặc mơ hồ; giữ mã chuỗi, loại dòng nhóm/tổng.
- Giá đơn, giá gộp độc lập VT/NC, tỷ lệ từ ô chọn; không tự đặt hệ số.
- Xem trước, chặn chồng nguồn, dữ liệu đổi, bảo vệ, công thức lỗi/vòng; ghi giá trị/công thức thật; bản sao lưu trên đĩa và khôi phục.
- Tài liệu XLSX/DOCX/PDF theo ba khu vực, lưu nguyên gốc/hash/phiên bản và trích chữ; nguồn giá có metadata và điều kiện so sánh.
- Quy tắc/mẫu riêng theo nghiệp vụ có phiên bản; biểu mẫu chỉnh bằng nút bấm; xuất XLSX/DOCX/PDF và JSON căn cứ.
- Shortcut, kiểm tra HTTPS/dịch vụ, script cài/gỡ; Python portable; Git bỏ qua dữ liệu nhạy cảm.

## Giới hạn cụ thể

Không có API AI trong bản này; không gửi dữ liệu ra ngoài hoặc giả lập AI. Chưa có OCR, so sánh kỹ thuật tự động/ghép ngữ nghĩa, trình dựng mọi công thức Excel, hộp Save As Windows riêng, sao lưu toàn bộ thư viện tự động. Công thức kiểm tra hỗ trợ tham chiếu A1 với số học và SUM/AVERAGE/MIN/MAX/COUNT/IF/ROUND/ABS; tên vùng, tham chiếu ngoài workbook, hàm động bị chặn. Cách tính thành tiền khác KL×(VT+NC) cần giữ nguyên file gốc và chưa được tính tự động. Quy tắc nghiệp vụ thực tế chưa được cung cấp nên khởi tạo chưa duyệt, cần người phụ trách hiệu chỉnh.

## Mã nguồn và build

`frontend/src`: UI, xử lý dòng, Office.js và runner nghiệm thu. `backend`: API, SQLite, chứng chỉ và xuất báo cáo. `manifests`: XML add-in. `scripts`: cài/khởi động/gỡ/đóng gói. `tests`: kiểm thử và workbook sideload riêng. Các thư mục dữ liệu theo yêu cầu được tạo tự động, không đưa vào Git.

Dev dùng Node 22+, pnpm và Python 3.12. `frontend/package.json` và lockfile xác định thư viện frontend; `backend/requirements.txt` xác định thư viện backend. Bản bàn giao đã có `frontend/dist`, `backend/vendor` và `runtime/python` nên người dùng không chạy build.

Khi phát triển, chạy tsc và Vite trong frontend, chạy `tests/core.test.ts` với tsx, chạy `tests/backend_test.py` với runtime Python. Backend kiểm thử cần httpx2 (chỉ phục vụ test). Không commit `config`, khóa/chứng chỉ riêng, SQLite, uploads, reports hay backups. Không cấu hình remote Git hoặc gửi mã/dữ liệu ra mạng trong bộ bàn giao.
