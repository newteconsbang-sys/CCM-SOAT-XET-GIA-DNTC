# Phiếu nghiệm thu CCM-AI 0.2

Chỉ dùng workbook thử hoặc bản sao không có dữ liệu thật. Kiểm thử tự động xử lý dữ liệu/dịch vụ không thay thế nghiệm thu Excel Desktop.

| Tình huống | Thao tác / kết quả mong đợi |
|---|---|
| Đọc workbook | Chọn vùng, lấy vùng: đúng sheet/địa chỉ, giá trị, STT chuỗi, công thức; danh sách sheet có sheet đang mở |
| Khác sheet/thứ tự | Báo giá: 1.10 trước 1.1; giá soát: đảo thứ tự, sheet khác; ghép đúng mô tả/ĐVT, không ghép theo vị trí |
| Giá đơn | KL 10, giá chào VT/NC 100/20, giá soát 90/18; chỉ sau xác nhận cách tính, tiền soát 1080, ảnh hưởng 120 |
| Giá gộp VT | Thành phần ống giá 80 × hệ số 1 và cách nhiệt 10 × 1, có nguồn; giá gộp 90; công thức liên kết PTDG |
| Giá gộp NC | Công tác lắp 12 × 1, công tác bọc 6 × 1, có nguồn; giá gộp 18 |
| Tỷ lệ | 5% chỉ trên hai ô giá trị 100 và 200; kết quả 15. Không tính trên toàn bảng; trùng ô bị chặn |
| Ghép mơ hồ | Lặp hai dòng STT/mô tả hoặc khác đơn vị; phải tự chọn dòng và xác nhận |
| Thiếu nguồn | Giá có số nhưng nguồn trống: chặn xuất/ghi; không có giá: để trống số, diễn giải Chưa đủ dữ liệu |
| STT là số | Cảnh báo mất số 0 cuối; cần xác minh chuỗi hiển thị, không tự đổi 1.1 thành 1.10 |
| Công thức lỗi | Giá nguồn #REF!, #DIV/0! hoặc ô nguồn thiếu số: chặn ghi; không biến thành số 0 |
| Vòng / tổng phụ | Cơ sở tỷ lệ trỏ vùng kết quả, vòng nguồn hoặc chọn tổng chứa ô con cũng được chọn: chặn xem trước |
| Bảo vệ / chồng nguồn | Chọn vùng đích chồng một nguồn hoặc sheet bảo vệ: chặn |
| Vùng có dữ liệu | Chặn lần đầu; đánh dấu cho phép thay, xem trước lại; chỉ Áp dụng mới ghi |
| Dữ liệu thay đổi | Sửa nguồn sau lấy vùng hoặc đích sau xem trước: chặn và yêu cầu lấy/xem lại |
| Khôi phục | Sau ghi đè, chọn bản sao và xác nhận: khôi phục giá trị, công thức, định dạng số; sheet mới do lần ghi tạo được xóa |
| Tính khi đóng add-in | Sau ghi, đóng task pane/dịch vụ; thay giá nguồn trong workbook: công thức Excel thông thường vẫn tính |
| Báo cáo | Giá chào, soát, chênh, ảnh hưởng, nguồn/lý do khớp giá đã xác nhận. Giá soát 0: % trống. Xuất Word/PDF giữ tiếng Việt |
| Mẫu | Chọn/bỏ/đổi thứ tự cột/mục và lưu; nghiệp vụ khác dùng phiên bản riêng |
| Không có AI | Chọn vùng, tính toán, sửa giá, xuất kết quả vẫn hoạt động; không tuyên bố có kết quả AI |

Ghi ngày, người kiểm tra, phiên bản Excel và kết quả từng tình huống. Các sheet TEST và báo cáo KIỂM THỬ đều là dữ liệu minh họa.


AI cục bộ: chạy thử minh họa, kiểm tra thiếu nguồn; chọn nguồn và xem trước nội dung; xác nhận phân tích; thử dùng/từ chối gợi ý. Gợi ý phải còn trạng thái Vàng và chưa xác nhận; workbook chỉ thay đổi sau xem trước và Áp dụng. Đổi dòng/nguồn trong lúc AI chạy phải chặn gợi ý cũ. Rút kết nối Ollama: vẫn thao tác soát xét thủ công được.
