# Kiểm tra CCM-AI 0.5 ngày 01/10/2026

Đã sao lưu và cập nhật bản cài D:\00.Setupgpt\SoatXetGia. Dịch vụ HTTPS phản hồi /api/health: CCM-AI, version 0.5.0, đúng thư mục cài. Chứng chỉ localhost còn hạn đến 01/10/2027 và đã được tin cậy.

Đã cấu hình CCM-AI-LocalService trong HKCU Run để chạy dịch vụ khi đăng nhập, không mở Excel hoặc Ollama. Chưa kiểm tra bằng đăng xuất/đăng nhập thực tế.

TypeScript kiểm tra thành công; Vite build thành công. 11 kiểm thử frontend, 4 kiểm thử thư viện/bridge, 5 kiểm thử backend đều đạt. Kiểm thử bridge dùng thư mục tạm, xác minh add-in lưu thư viện rồi endpoint đọc đúng phiên bản, giữ giá 0 và trống, không đọc lẫn nghiệp vụ, không chấp nhận thư viện hỏng. Không đưa giá minh họa vào thư viện thật.

Giao diện web bản cài tải được, có nút gửi thư viện cho skill; nút chặn thao tác workbook khi mở bằng trình duyệt không có Excel. Đã lưu hướng dẫn bridge vào skill cá nhân soatxetgia-ccm trong tài khoản ChatGPT (skill id 6abe2821384c8191ad026340e9c676c6). Đã thấy ChatGPT Excel có khung chat và phiên workbook kết nối.

ĐNTC, EW, NCC, NTP ở bản cài đều 0 đầu mục. Endpoint snapshot trả 409 với thông báo nạp thư viện, không tạo giá. Chưa xác minh việc add-in ghi sheet bridge hoặc skill đọc/điền giá trên Excel thật: công cụ đọc danh sách sheet hết thời gian chờ; công cụ chụp giao diện Excel báo FrameArrived/window capture timed out. Chưa sửa ô giá nào trong workbook tài chính của người dùng trong lần kiểm tra này.

Sau khi nạp dữ liệu thật: gửi thư viện từ CCM-AI, đọc metadata và đủ dòng bằng ChatGPT Excel, xem trước ô đích, giữ Lô, ghi theo cột được chọn rồi đọc lại. Khi thư viện đổi cần gửi lại snapshot.
