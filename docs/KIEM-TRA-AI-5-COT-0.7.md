# CCM-AI 0.7.0 — nhận diện thư viện 5 cột

Đã cập nhật bản cài D:\00.Setupgpt\SoatXetGia, HTTPS localhost:3443. Giữ dữ liệu và cấu hình hiện hành. Sao lưu trước cập nhật: backups/ai-5-cot-0.7-20261002/before.

- 22 kiểm tra backend đạt: nhận diện tiêu đề nhiều dòng, loại STT/thành tiền, AI không hợp lệ, định dạng giá/khối lượng riêng, khối lượng 0, PDF/XLSX, lưu có phiên bản, skill đọc khối lượng, JSON cầu nối và Excel đúng 5 cột.
- 11 kiểm tra frontend đạt; TypeScript và Vite build đạt.
- Đã gọi Qwen qwen3:4b trên máy thật, không dùng cloud. Web hiển thị AI cục bộ khi mô hình trả ánh xạ hợp lệ; nếu lỗi báo dùng nhận diện tiêu đề.
- PDF ĐNTC 18 trang / 18 bảng: AI tự chọn bảng gộp, dòng tiêu đề 2; hạng mục B, đơn vị C, khối lượng D, vật tư F, nhân công G. Nhận diện 606 đầu mục có giá số; 111 dòng bỏ qua (gồm dòng nhóm/thiếu dữ liệu/không có giá số).
- 49 ô giá tỷ lệ/biểu thức chưa quy đổi được liệt kê riêng. Không suy ra 5%VT hay VT VTP/3; đầu mục còn giá số giữ ô thiếu trống và lưu biểu thức trong điều kiện nguồn. Đầu mục không còn giá số được bỏ qua; bản PDF gốc vẫn giữ nguyên.
- Excel xuất thực tế qua API local có xác thực: NTC-DNTC-Thu-vien-5-cot.xlsx, 606 dòng / đúng 5 cột. Đọc lại và so sánh toàn bộ giá trị với dữ liệu nguồn; không có ô công thức, SHA256 nguồn nằm trong thuộc tính workbook.
- Bảng đã kiểm tra ở khung trình duyệt hẹp 664px, đủ 5 cột, không tràn trang. Có liên kết tải trực tiếp sau tạo file. Trình duyệt trong Codex không trả sự kiện tải / chưa thấy file trong Downloads; bản Excel đã xác minh được lưu tại D:\SoatXetGia và bàn giao trực tiếp.
- PDF ĐNTC hiện để ở bước xem trước, chưa thay thư viện thật. Người dùng đối chiếu PDF, đánh dấu đã kiểm tra rồi bấm Lưu 606 đầu mục vào ĐNTC để skill đọc bản hiện hành.
- Khi soát CCM, không ghi đè khối lượng và không áp công thức Lô. PDF ảnh scan chưa có OCR.
