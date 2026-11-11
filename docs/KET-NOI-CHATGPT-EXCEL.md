# CCM-AI và skill trong ChatGPT Excel

Bản 0.5 dùng một sheet thư viện trong chính workbook làm cầu nối dữ liệu. Skill không cần gọi dịch vụ localhost từ cloud.

1. Mở CCM-AI trong Excel, chọn ĐNTC, EW, NCC hoặc NTP.
2. Trong mục thư viện riêng, chọn vùng giá, ánh xạ cột rồi lưu thư viện.
3. Bấm **Gửi thư viện … cho skill ChatGPT**. Add-in tạo sheet CCM_AI_<nghiệp vụ>_<mã>, đọc lại để xác minh và lưu bản sao thao tác. Các sheet hiện có được giữ nguyên.
4. Mở ChatGPT Excel, chọn skill **soatxetgia ccm**. Nói nghiệp vụ, sheet CCM, vùng dòng và hai cột giá muốn điền. Ví dụ: “Dùng thư viện ĐNTC vừa gửi từ CCM-AI để soát vùng CCM đang chọn. Giá vật tư vào cột P, nhân công vào Q. Xem trước, giữ nguyên Lô.”
5. Thư viện đổi thì bấm gửi lại. Mỗi sheet là một bản sao có phiên bản, thời điểm gửi, số dòng và SHA256 nguồn; không tự cập nhật.

Skill chỉ lấy giá từ đúng nghiệp vụ, giữ giá 0, bỏ giá thiếu/không khớp và giữ nguyên Lô/công thức. Có sheet mà phiên ChatGPT không có công cụ đọc workbook thì skill phải báo chưa kết nối; không báo đã lấy giá.

Dịch vụ cục bộ tự khởi động khi đăng nhập Windows, không mở Excel/Ollama. Có thể mở lại bằng shortcut **Khởi động CCM-AI** nếu dịch vụ dừng. Khi cập nhật dịch vụ, tải lại bảng CCM-AI; bản 0.5 tự xin lại phiên truy cập khi token cũ hết hiệu lực.

Dữ liệu được ghi vào workbook người dùng. ChatGPT sẽ xử lý phần workbook được đọc theo tài khoản đang sử dụng; chọn thư viện cần gửi. Không có thư viện nào được công khai qua Internet.

Nguồn API Excel: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-worksheets
