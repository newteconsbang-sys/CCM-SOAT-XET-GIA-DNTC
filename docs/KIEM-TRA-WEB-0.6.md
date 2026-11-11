# Kiểm tra và cài đặt web local 0.6 — 01/10/2026

- Bản cài thực tế: `D:\00.Setupgpt\SoatXetGia`; HTTPS `/api/health` trả app CCM-AI, version 0.6.0, đúng installPath.
- Desktop có shortcut **CCM-AI Soat gia (Web)**. Mở web không bắt buộc Excel/Ollama. Add-in cũ dùng đường `/?excel=1`.
- TypeScript strict và Vite build đạt; 11 kiểm thử frontend, 5 backend, 4 cầu nối thư viện và 5 luồng web đạt.
- Luồng web kiểm thử trong thư mục tạm: XLSX/CSV/JSON → nạp → ánh xạ → xem trước → xác nhận → lưu phiên bản → helper đọc/dò. Kiểm tra giá 0, giá thiếu, dòng Lô, không lẫn nghiệp vụ, công thức nguồn bị chặn, file gốc đổi hash bị chặn, phiên bản đổi trước khi dò bị chặn.
- Trình duyệt thực tế đã nạp XLSX mẫu và xem trước 3 đầu mục đúng giá 0/thiếu/Lô; chưa xác nhận lưu dữ liệu mẫu vào thư viện đang dùng. Kiểm thử lưu dùng dữ liệu thử tách biệt.
- Helper đã cài vào skill cá nhân trong Codex; chạy `--action list` đọc đúng cả bốn nghiệp vụ hiện chưa nạp. Trình kiểm tra skill chính thức báo `Skill is valid!`.
- Skill cá nhân trong tài khoản ChatGPT đã lưu hướng dẫn đọc web local trong Codex và báo rõ hạn chế của phiên cloud.
- Backup mã trước nâng cấp: `D:\SoatXetGia\backups\web-0.6-20261001-172603`.

Chưa soát/ghi giá vào hồ sơ tài chính thật hoặc kiểm chứng ghi workbook trực tiếp bằng Office.js. Thư viện thực tế cần người dùng nạp. Việc cập nhật skill tài khoản không tạo kết nối tự động từ khung ChatGPT Excel đến ổ D.
