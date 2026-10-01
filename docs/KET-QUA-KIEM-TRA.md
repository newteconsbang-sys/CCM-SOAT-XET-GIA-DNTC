# Kết quả kiểm tra – 01/10/2026

## Đã chạy thật ở lớp mã và dịch vụ

- TypeScript strict: đạt; Vite build React: đạt.
- 7/7 kiểm thử xử lý dữ liệu: STT chuỗi 1.10/1.1; hai sheet đảo thứ tự; ghép trùng/mơ hồ/khác đơn vị; dòng nhóm/tổng và lỗi nguồn; giá gộp vật tư/nhân công; cơ sở tỷ lệ; mẫu số 0; chồng nguồn; đối soát tiền soát/ảnh hưởng; chặn xác nhận/nguồn/chứng cứ thiếu.
- 5/5 kiểm thử FastAPI: phiên bản quy tắc; nguồn giá bắt buộc; sao lưu trên đĩa và không ghi đè backup; giới hạn origin/session; xuất XLSX/DOCX/PDF.
- Báo cáo kiểm thử dùng số liệu minh họa: STT `1.10`, giá chào 100/20, giá soát 90/18, KL 10, tiền soát 1080, ảnh hưởng 120. Đọc lại XLSX đã xuất đúng STT và số liệu; ba định dạng đã tạo thật. Chưa nghiệm thu mọi bố cục Word/PDF trên hồ sơ dài thực tế.
- Runtime Python portable đã được chạy và kiểm tra import thư viện.
- Dịch vụ đã chạy HTTPS `https://localhost:3443/api/health`; xác minh bằng chính chứng chỉ CCM đã tạo, phản hồi app CCM-AI, AI disabled. Đây không phải bằng chứng Excel/WebView đã tin cậy chứng chỉ.

## Cài đặt trên máy

Office phát hiện: Microsoft 365 O365ProPlusRetail, x64, 16.0.20430.20118. Manifest đã đăng ký vào HKCU Office 16.0 WEF Developer cho đúng ID CCM; shortcut Khởi động CCM-AI đã tạo trên Desktop và trong D:\SoatXetGia, cấu hình trỏ runtime portable. Chứng chỉ localhost đã tạo nhưng bước tin cậy Windows chưa được xác nhận tại thời điểm lập báo cáo.

## Chưa được xác nhận

Chưa có kết quả Office.js đọc/ghi trong Excel Desktop thực tế: không tuyên bố tích hợp thành công. Workbook kiểm thử riêng và runner đã chuẩn bị; cần hoàn tất tin cậy chứng chỉ, mở workbook nghiệm thu trong Excel và kiểm tra file `local-officejs-acceptance.json`. Việc tự động chặn bảo vệ, tham chiếu vòng, ghi đè/khôi phục ở lớp Office.js cần được xác nhận bằng workbook đó và phiếu NGHIEM-THU.md.

Không có API AI, OCR hoặc bộ quy tắc nghiệp vụ thực tế đã duyệt. Bản hiện tại cần nghiệm thu cá nhân trước khi dùng hồ sơ thật; chưa được coi hoàn tất toàn bộ yêu cầu đầu vào.
