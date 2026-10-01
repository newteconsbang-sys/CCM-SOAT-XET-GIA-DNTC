# Kết quả kiểm tra – 01/10/2026

## Đã chạy thật ở lớp mã và dịch vụ

- TypeScript strict: đạt; Vite build React: đạt.
- 7/7 kiểm thử xử lý dữ liệu: STT chuỗi 1.10/1.1; hai sheet đảo thứ tự; ghép trùng/mơ hồ/khác đơn vị; dòng nhóm/tổng và lỗi nguồn; giá gộp vật tư/nhân công; cơ sở tỷ lệ; mẫu số 0; chồng nguồn; đối soát tiền soát/ảnh hưởng; chặn xác nhận/nguồn/chứng cứ thiếu.
- 5/5 kiểm thử FastAPI: phiên bản quy tắc; nguồn giá bắt buộc; sao lưu trên đĩa và không ghi đè backup; giới hạn origin/session; xuất XLSX/DOCX/PDF.
- Báo cáo kiểm thử dùng số liệu minh họa: STT `1.10`, giá chào 100/20, giá soát 90/18, KL 10, tiền soát 1080, ảnh hưởng 120. Đọc lại XLSX đã xuất đúng STT và số liệu; ba định dạng đã tạo thật. Chưa nghiệm thu mọi bố cục Word/PDF trên hồ sơ dài thực tế.
- Runtime Python portable đã được chạy và kiểm tra import thư viện.
- Dịch vụ đã chạy HTTPS `https://localhost:3443/api/health`; xác minh bằng chính chứng chỉ CCM đã tạo, phản hồi app CCM-AI, AI disabled. Đây không phải bằng chứng Excel/WebView đã tin cậy chứng chỉ.

## Cài đặt trên máy

Office phát hiện: Microsoft 365 O365ProPlusRetail, x64, 16.0.20430.20118. Bộ bàn giao đã đặt vào đúng `D:\00.Setupgpt\SoatXetGia` khi thư mục này chưa tồn tại, không ghi đè dữ liệu cũ. Manifest đã đăng ký vào HKCU Office 16.0 WEF Developer cho đúng ID CCM, đường dẫn manifest ở thư mục đó; shortcut Khởi động CCM-AI trên Desktop trỏ runtime portable của bản cài. Mã nguồn và Git phát triển giữ thêm tại `D:\SoatXetGia`. Chứng chỉ localhost đã tạo nhưng bước tin cậy Windows chưa được xác nhận tại thời điểm lập báo cáo.

## Chưa được xác nhận

Chưa có kết quả Office.js đọc/ghi trong Excel Desktop thực tế: không tuyên bố tích hợp thành công. Workbook kiểm thử riêng và runner đã chuẩn bị; cần hoàn tất tin cậy chứng chỉ, mở workbook nghiệm thu trong Excel và kiểm tra file `local-officejs-acceptance.json`. Việc tự động chặn bảo vệ, tham chiếu vòng, ghi đè/khôi phục ở lớp Office.js cần được xác nhận bằng workbook đó và phiếu NGHIEM-THU.md.

Tại bản 0.1 chưa có API AI, OCR hoặc bộ quy tắc nghiệp vụ thực tế đã duyệt. AI cục bộ được bổ sung trong bản 0.2 bên dưới. Bản hiện tại cần nghiệm thu cá nhân trước khi dùng hồ sơ thật; chưa được coi hoàn tất toàn bộ yêu cầu đầu vào.

## Cập nhật cài đặt trên máy

Windows đã tin cậy đúng chứng chỉ của bản cài ở `D:\00.Setupgpt\SoatXetGia`, thumbprint `B17825672AF5F5E733AF8B24CED96B9D85E455E1`. Dịch vụ được khởi động bằng script cài và kiểm tra HTTPS thành công, phản hồi đúng installPath. Excel đã được khởi động và đã gửi lệnh mở workbook nghiệm thu riêng. Công cụ đọc cửa sổ Excel lỗi `FrameArrived timed out` rồi `window capture timed out` sau một lần phục hồi; không thao tác tiếp dựa trên tọa độ cũ. Chưa nhận được file kết quả Office.js, nên trạng thái nghiệm thu đọc/ghi Excel vẫn chưa xác nhận. Người dùng có thể tiếp tục bằng Home → Add-ins → Advanced → My Add-ins → Developer Add-ins → CCM-AI → Add.

## Bản 0.2 — AI cục bộ (2026-10-01)

Đã cài Ollama v0.35.0 CPU với chữ ký Ollama Inc. hợp lệ và Qwen3 4B Q4_K_M (2.5 GB) tại thư mục cài CCM-AI. Model lưu trong data/models; cấu hình enabled=true, model=qwen3:4b. Log tiến trình xác nhận OLLAMA_NO_CLOUD=true và Ollama cloud disabled: true. Backend HTTPS đã chạy phiên bản 0.2.0 tại đúng installPath.

- TypeScript strict và Vite build đạt; 7 kiểm thử core, 5 kiểm thử backend và 8 kiểm thử AI đạt. Kiểm thử AI tự động dùng mock transport, không phải bằng chứng suy luận thật.
- Đã gọi suy luận thật bằng Qwen3 4B trên máy. Mẫu không có nguồn trả Chưa đủ dữ liệu, lần đầu khoảng 41.6 giây. Mẫu hai dòng cùng mô tả/ĐVT gợi ý ref:0, giữ đúng đơn giá vật tư 20.000 từ nguồn so với giá chào 22.000, không tự tạo nguồn giá thư viện, trạng thái vàng. Phản hồi từ chối mẫu kiểm thử đã lưu SQLite, chưa ghi workbook.
- Màn hình trình duyệt đã nhận Ollama đang chạy, mô hình qwen3:4b. Đây là xác nhận giao diện web và API, chưa phải đọc/ghi Excel Desktop.
- Cơ chế schema động chỉ cho model chọn các ID trong bản xem trước; backend kiểm tra lại nguồn/JSON. Ngữ cảnh xem trước hết hạn hoặc đã dùng bị chặn; đổi mô hình/quy tắc cần chuẩn bị lại. Giá được lấy lại từ nguồn, chưa phải AI xác minh tính tương đương.

Giới hạn: AI vẫn có thể diễn giải thiếu chính xác, kể cả cảnh báo một điều kiện đã ghi trong mô tả. Người dùng phải đối chiếu thông số, VAT và phạm vi trước khi chấp nhận. AI chưa tự phân rã giá gộp, chọn hệ số hoặc huấn luyện từ phản hồi. Nghiệm thu Office.js trong Excel Desktop vẫn chưa xác nhận; các kết quả trên dùng dữ liệu minh họa.

## Bản 0.3 — Ghi nhớ nghiệp vụ (2026-10-01)

TypeScript/Vite đạt; 13 kiểm thử API AI đạt, gồm lưu ví dụ vào SQLite và lấy lại sau khi tạo router mới, duyệt bắt buộc, tách nghiệp vụ/ĐVT, bỏ giá lịch sử khỏi ngữ cảnh, tắt tra cứu và thu hồi. 5 kiểm thử backend hồi quy đạt. Kiểm thử thật trên bản đã cài đã lưu ví dụ minh họa, lấy vào approvedExamples, gọi Ollama và thu hồi ví dụ sau kiểm thử. Giá hiện tại thiếu vẫn để null. Mô hình thử nghiệm chỉ trả Chưa đủ dữ liệu, chưa diễn giải bài học; ứng dụng vẫn hiển thị bài học đã truy xuất ở recalledExamples. Không tuyên bố AI luôn áp dụng đúng ví dụ. Giao diện web bản 0.3 có mục Ghi nhớ nghiệp vụ đã duyệt; đọc/ghi Excel Desktop vẫn cần nghiệm thu.
