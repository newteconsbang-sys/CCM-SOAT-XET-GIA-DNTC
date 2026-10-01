# Sử dụng CCM-AI – Soát xét giá

## Luồng làm việc

1. Bấm **Khởi động CCM-AI**, mở workbook báo giá, bấm tab/nút CCM-AI.
2. Chọn nghiệp vụ Điện nước ĐNTC / EW / NTP / NCC và nhập dự án.
3. Chọn vùng báo giá bằng chuột, kể cả hàng tiêu đề nếu có; bấm **Lấy vùng Excel đang chọn** ở Báo giá NCC/NTP. Làm tương tự với vùng Giá soát xét; hai vùng có thể ở sheet khác nhau.
4. Kiểm tra ánh xạ từng cột. Chọn **Không có** cho trường thiếu. ĐVT/KL nằm ngoài vùng: chọn cột đơn có cùng số dòng, kể cả tiêu đề, bấm lấy cột ngoài vùng tương ứng. Nếu dùng chung, chọn cùng cột đó cho cả hai vùng. Xác nhận từng ánh xạ. STT nên lưu dạng văn bản để giữ `1.10` khác `1.1`; khi STT đang là số cần xác minh chuỗi hiển thị.
5. Có thể lưu tên mẫu ánh xạ để dùng lần sau; mẫu chỉ giữ vai trò cột, không khóa địa chỉ hoặc sheet. Luôn xác nhận ánh xạ sau khi lấy vùng mới.
6. Nạp tài liệu vào đúng khu vực Hồ sơ cần soát / Nguồn tham chiếu / Hướng dẫn AI. PDF quét chưa hỗ trợ OCR; tài liệu Excel/Word/PDF được lưu nguyên gốc và trích chữ để tra cứu. Nạp tệp không tự biến thành thư viện giá đã xác nhận: nhập giá và nguồn có căn cứ trong Quản lý thư viện giá.
7. Chỉnh và lưu quy tắc cho nghiệp vụ. Quy tắc khởi tạo chưa được duyệt; đây là các kiểm tra chung, không phải quy định CCM/BRAVO đã được cung cấp. Người phụ trách bổ sung/duyệt theo nghiệp vụ thực tế. Mỗi nghiệp vụ có phiên bản riêng.
8. Bấm **Bắt đầu soát xét**. Hệ thống ghép chắc chắn khi STT và toàn bộ mô tả khớp, đơn vị phù hợp và chỉ có một ứng viên; không tự ghép theo vị trí. Chọn lại dòng đối chiếu và xác nhận đối với dòng trùng/thiếu/mơ hồ. Xác minh hãng/model/thông số, đơn vị, VAT, vận chuyển, phạm vi và thời điểm; không tự lấy giá thấp nhất.
9. Chọn dòng, mở tab Vật tư hoặc Nhân công. Có thể chọn một ô thuộc dòng báo giá rồi bấm **Mở Phân tích đơn giá theo ô chọn**. Bảng phân tích thực tế nằm trong task pane; không có popup giả lập trong ô Excel.

## Giá đơn, giá gộp và tỷ lệ

- **Giá đơn:** nhập giá được xác nhận và nguồn. Chọn thư viện giá bằng danh sách nếu phù hợp; nguồn và điều kiện đi kèm. Giá thư viện chỉ tham khảo cho đến khi xác minh tương đương.
- **Giá gộp:** thêm từng thành phần, ĐVT, hệ số và đơn giá có nguồn. Giá đóng góp = hệ số × đơn giá; giá gộp = tổng đóng góp. Hỗ trợ độc lập cho vật tư và nhân công. Hệ số để trống nếu thiếu căn cứ; hệ thống không tự điền 1 hoặc đoán thành phần.
- **Tỷ lệ:** nhập % và chỉ chọn ô số làm cơ sở; bấm thêm vùng. Có thể thêm các vùng khác sheet. Bỏ ô thừa, không chọn cả tổng phụ và các dòng con. Hệ thống kiểm tra trùng ô và lần theo công thức A1 thông dụng để chặn phụ thuộc lồng nhau/vòng tham chiếu. Công thức động, tên vùng, hàm chưa hỗ trợ kiểm tra hoặc đồ thị quá lớn bị chặn để bạn dùng nguồn đơn giản đã xác nhận.
- Công thức nguồn được đọc/hiển thị và giữ nguyên ở vùng gốc. Kết quả dùng tham chiếu có tên sheet và địa chỉ tuyệt đối đến giá nguồn hoặc sheet PTDG; không sao chép công thức tương đối sang vị trí mới.
- Chưa có trình biên tập tùy ý mọi công thức Excel: hỗ trợ công thức tỷ lệ từ ô chọn, giá gộp và các công thức chênh lệch/thành tiền đã xác nhận.

## Đánh giá và thành tiền

Xanh cần quy tắc đã được duyệt, các kiểm tra của dòng đã đánh dấu và có đủ giá/căn cứ. Vàng dành cho nội dung chưa đủ dữ liệu hoặc cần kiểm tra. Đỏ phải nhập chứng cứ lỗi. Giá cao hơn đối chiếu không tự là lỗi. Cần nhập lý do cho từng dòng.

Mặc định không tính thành tiền. Chỉ chọn **Tôi xác nhận: KL × (ĐG vật tư + ĐG nhân công)** khi đúng cách tính của file. Với cách tính khác, để trống thành tiền/ảnh hưởng; chưa có trình dựng cách tính khác trong phiên bản này. Phần trăm chênh lệch = (giá chào − giá soát) / giá soát × 100; không tính khi mẫu số 0. Ảnh hưởng = thành tiền chào trong file − thành tiền soát đã xác nhận.

## Ghi kết quả, sao lưu và khôi phục

1. Chọn sheet mới hoặc sheet đang có, ô bắt đầu và Ghi giá trị / Chèn công thức Excel.
2. **Xem trước** để đọc địa chỉ, giá trị, công thức sẽ ghi. Sheet kết quả có 15 cột; có hàng tiêu đề. Sheet PTDG tạo riêng khi dùng giá gộp với chế độ công thức. Dòng thiếu giá để trống kết quả số và giữ chữ Chưa đủ dữ liệu trong diễn giải.
3. Hệ thống chặn vùng chồng nguồn, sheet được bảo vệ, dữ liệu nguồn đổi sau khi lấy vùng, tham chiếu lỗi/vòng và vùng đích đã đổi sau xem trước. Nếu vùng đích có dữ liệu, cần đánh dấu cho phép thay dữ liệu rồi xem trước lại.
4. Chỉ **Áp dụng** mới ghi. Bản sao trước thay đổi lưu ở `backups` trên ổ đĩa, kèm workbook ID, sheet, địa chỉ, giá trị, công thức, định dạng số. Không chỉ dựa vào Ctrl+Z. Ghi thất bại cố khôi phục các vùng đã thay; nếu mất kết nối hoặc workbook đóng giữa giao dịch, dùng bản sao trạng thái prepared để khôi phục thủ công.
5. **Khôi phục bản sao:** chọn bản sao trong đúng workbook, đọc phạm vi, đánh dấu xác nhận rồi khôi phục. Vùng cũ được khôi phục; sheet do lần ghi đó tạo sẽ bị xóa, gồm các chỉnh sửa bạn đã thêm vào sheet sau đó. Bản sao không phục hồi cấu trúc workbook nếu bạn đã đổi tên/xóa sheet gốc.

## Xuất báo cáo

Xuất báo cáo tách riêng với ghi workbook. Nhập tiêu đề; chọn toàn bộ hoặc chỉ dòng cần lưu ý; xem trước. Trong tùy chỉnh mẫu, chọn/bỏ cột, bấm ↑/↓ đổi thứ tự và nhập thứ tự mục, lưu phiên bản. Bấm Xuất Excel/Word/PDF. Dữ liệu xuất lấy từ các giá và cách tính đã xác nhận trong task pane; không tự đọc lại giá bạn sửa trực tiếp trên sheet kết quả. Bản báo cáo và JSON căn cứ được lưu ở `reports`.

File tải về qua cơ chế tải của Excel/WebView. Muốn tự chọn vị trí mỗi lần, bật tùy chọn hỏi vị trí tải của trình duyệt hoặc lấy file trong reports rồi Save As. Chưa có hộp thoại Save As Windows riêng do Office.js task pane không cung cấp trực tiếp đường dẫn hệ thống tùy ý.

## AI và dữ liệu

Chưa cấu hình AI và chưa tích hợp nhà cung cấp API trong bản này; không có kết quả AI giả lập, không gửi hồ sơ ra ngoài và không có phí API AI. Các chức năng hiện tại dùng logic xác định và xác nhận của người dùng. ChatGPT không được coi là khóa API. Muốn thêm AI cần lựa chọn nhà cung cấp, cấu hình backend, cơ chế xin phép gửi dữ liệu và phí riêng; không đặt khóa ở frontend hoặc Git.

Tài liệu có nội dung hướng dẫn chỉ là dữ liệu tham khảo; không được thực thi như lệnh. SQLite lưu phiên bản nguồn giá, quy tắc, mẫu, ánh xạ và tài liệu. Thư mục hồ sơ, báo cáo, bản sao, cấu hình riêng, chứng chỉ/khóa đều bị loại khỏi Git. Chưa có tự động sao lưu toàn bộ thư viện sang ổ khác; nên sao chép các thư mục dữ liệu khi dịch vụ đã dừng để bảo toàn SQLite.
