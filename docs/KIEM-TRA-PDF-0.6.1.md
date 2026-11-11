# Nạp PDF và chuyển Excel — bản 0.6.1, 02/10/2026

Đã bổ sung chọn PDF tại web local, tách bảng chữ theo trang, gộp bảng có toàn bộ tiêu đề giống nhau, tải Excel chứa bảng nguyên gốc và nguồn trang, chọn định dạng giá, xem trước rồi lưu vào nghiệp vụ đã chọn. Nguồn thư viện giữ PDF gốc, hash và trang/bảng/dòng. Chọn lại nghiệp vụ đang dùng không xóa file nháp.

Kiểm thử tự động đạt: 5 kiểm thử PDF + 5 kiểm thử nạp web; 11 kiểm thử frontend; TypeScript strict và Vite build. Bao phủ PDF hai trang, gộp tiêu đề, giá dấu chấm/phẩy, giá 0/trống, giữ Lô, thiếu xác nhận PDF bị chặn, skill đọc/dò giá sau lưu, tải Excel đọc lại đúng chữ và nguồn trang, chuỗi công thức giữ dưới dạng chữ và bị chặn khi dùng làm giá, mật khẩu, trang chưa đọc được và bảng chuyển bị sửa hash.

Thư viện kiểm thử lưu trong thư mục tạm riêng. Chưa lưu dữ liệu minh họa vào thư viện dùng thật. Luồng viết workbook tài chính thật chưa được nghiệm thu.

Bản đã cài tại `D:\00.Setupgpt\SoatXetGia`: HTTPS health xác nhận 0.6.1, parser đọc PDF hai trang thành hai bảng. Trình duyệt đã nạp PDF thử, chọn dấu chấm hàng nghìn và xem trước 4 đầu mục đúng giá 32.800, giá 0, giá thiếu và nguồn trang 2. Nút **Tải Excel từ PDF** đã tạo file thực trong Downloads; đọc lại đúng 3 sheet, giá chữ gốc và nguồn trang PDF. Bốn thư viện thật vẫn chưa nạp, không lưu giá thử. Mã cũ được sao lưu tại `D:\SoatXetGia\backups\pdf-0.6.1-20261002\before`.

Giới hạn: 25 MB/PDF, 50 trang, 30 bảng, 10.000 dòng/100 cột mỗi bảng, 500.000 ô. Chưa có OCR cho PDF ảnh scan. Cảnh báo nêu trang chưa đọc được và bảng suy ra từ chữ; người dùng cần kiểm tra cột, dòng, thương hiệu và điều kiện giá trước lưu. XLSX chuyển đổi là dữ liệu chữ nguyên gốc; chuẩn hóa giá chỉ diễn ra khi xem trước với định dạng do người dùng chọn.
