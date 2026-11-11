# Web local CCM-AI và skill soát giá

## Mở ứng dụng

Trên máy đã cài: mở shortcut Desktop **CCM-AI Soat gia (Web)** hoặc chạy `D:\00.Setupgpt\SoatXetGia\Mo-web-CCM-AI.cmd`. Script khởi động dịch vụ rồi mở **https://localhost:3443/**. Không cần mở Excel hoặc mô hình Ollama để nạp thư viện. Web prepareccm tại cổng 8000 được giữ nguyên.

## Nạp thư viện

1. Chọn **ĐNTC, EW, NCC hoặc NTP** ở thanh bên. Mỗi nghiệp vụ có dữ liệu và phiên bản riêng.
2. Trong **Nạp thư viện**, chọn XLSX, PDF, CSV hoặc JSON, tối đa 25 MB. AI Qwen cục bộ tự nhận diện bảng, dòng tiêu đề và 5 cột: **Hạng mục, Đơn vị, Khối lượng, Đơn giá vật tư, Đơn giá nhân công**. Lần đầu khởi động mô hình có thể mất khoảng một phút.
3. Kết quả xem trước tự hiển thị. Kiểm tra hạng mục, đơn vị, khối lượng và giá rồi bấm **Lưu … đầu mục** vào nghiệp vụ đã chọn. Với PDF phải đối chiếu bản gốc và đánh dấu đã kiểm tra trước lưu.
4. Nếu nhận diện chưa đúng, mở **Kiểm tra / chỉnh cột** để chọn sheet, dòng tiêu đề, phạm vi, 5 cột và định dạng số. Có nút **Nhận diện lại bằng AI**. AI chưa sẵn sàng thì giao diện báo rõ đang nhận diện theo tiêu đề.
5. Bấm **Tải Excel 5 cột** để tải bảng đã nhận diện thành số. Trong **Giá đã lưu**, tìm đầu mục hoặc tải JSON để dùng trong phiên chat cần file đính kèm.

Cần hạng mục, đơn vị và ít nhất một loại giá. Khối lượng thiếu giữ trống, không tự điền 1; số 0 hợp lệ. Không lấy STT hoặc thành tiền làm khối lượng/đơn giá. Giá trống là thiếu giá; ô giá/khối lượng công thức cần chuyển thành giá trị. Định dạng dấu phân cách được nhận diện riêng cho giá và khối lượng: PDF có thể có khối lượng `45.00` nhưng giá `32.800`; phải kiểm tra khi có cảnh báo chưa thống nhất. Dòng không đủ dữ liệu được bỏ qua và báo khi xem trước.

Giá tỷ lệ / biểu thức nguồn như `5%VT`, `VT VTP/3` chưa được tính: ô giá tương ứng giữ trống, đầu mục không có giá số được bỏ qua; xem trước báo số ô và dòng cần thiết lập công thức. Bản gốc giữ nguyên; điều kiện của đầu mục còn giá số lưu cả biểu thức chưa quy đổi.

AI chỉ xác định cột; số liệu được đọc từ file gốc. Mỗi file được giữ nguyên cùng SHA256; nguồn giá lưu tên file, sheet và dòng. Thư viện hiện hành chỉ đổi khi bấm lưu; lưu lần nữa thay bản hiện hành của đúng nghiệp vụ, không cộng dồn tự động. Khối lượng trong thư viện là dữ liệu nguồn, không dùng để ghi đè khối lượng CCM khi soát giá.

## Nạp và chuyển PDF

PDF có bảng chữ đọc được, tối đa 50 trang và 30 bảng. Web tách bảng theo trang, gộp các bảng cùng toàn bộ dòng tiêu đề và giữ bảng từng trang. AI chọn bảng phù hợp; có thể chỉnh lại trong phần kiểm tra cột.

**Tải Excel 5 cột** chuyển bảng đã nhận diện thành XLSX với đúng 5 cột và giá/khối lượng dạng số. Tên file và SHA256 nằm trong thuộc tính workbook. Giá lưu còn kèm trang/bảng/dòng PDF nguồn. Skill đọc thư viện đã lưu như nguồn Excel.

PDF có mật khẩu phải mở khóa trước. PDF ảnh scan chưa có OCR: web báo rõ và không tạo giá từ ảnh. Cảnh báo trang chưa tách được phải được kiểm tra; bảng đọc được không đại diện cho toàn bộ PDF. Dòng xuống dòng/ô gộp cần đối chiếu bản gốc trước khi dùng.

## Gọi skill trong Codex trên máy

Gửi file CCM hoặc chỉ rõ file trên máy, sheet, phạm vi dòng và các cột nguồn/đích. Ví dụ:

> Dùng $soatxetgia-ccm lấy thư viện NTP đã lưu trên web local CCM-AI để soát file CCM tôi gửi. Mô tả cột B, ĐVT cột D, thương hiệu cột C; điền vật tư vào cột tôi chọn và nhân công vào cột tôi chọn. Xem trước và giữ nguyên dòng Lô.

Thay hai cột đích bằng cột thực tế. Skill đọc thư viện từ `D:\00.Setupgpt\SoatXetGia\libraries` qua helper đã cài, xác nhận nghiệp vụ, phiên bản và nguồn trước khi dò. Dò mô tả/ĐVT/thương hiệu/thông số; nhiều nguồn hoặc không khớp thì giữ nguyên để kiểm tra. Giá thiếu giữ nguyên ô tương ứng; mọi dòng **Lô** và ô đang có công thức được giữ nguyên. Chưa áp công thức Lô.

Helper chỉ đọc và dò giá. Khi yêu cầu điền, phiên cần công cụ sửa file/Excel phù hợp; phải lưu bản sao, kiểm tra ô đích và đọc lại sau ghi. Chưa ghi thì skill báo rõ, không coi đề xuất là đã sửa workbook.

## Khung ChatGPT trong Excel

Để khung ChatGPT Excel đọc dữ liệu giá đã lưu, đưa bản snapshot vào chính workbook đang mở:

1. Mở workbook và khung ChatGPT Excel, đăng nhập. Trong web local chọn nghiệp vụ, mở **Dùng với skill** và bấm **Sao chép lệnh kết nối Excel**.
2. Dán lệnh vào **Codex trên máy**. Codex xác nhận đúng workbook kết nối, đọc đủ thư viện hiện hành và tạo sheet nguồn `CCM_AI_DNTC_…` / EW / NCC / NTP. Đọc lại dữ liệu và kiểm tra các sheet hồ sơ được giữ nguyên trước báo đã kết nối.
3. Trong khung ChatGPT Excel, gọi `soatxetgia-ccm`, chỉ rõ sheet nguồn, sheet hồ sơ và hai cột đích. Skill đọc metadata, đủ số dòng và nguồn trước khi dò/điền giá.

Nút **Gửi thư viện … cho skill ChatGPT** là nút của CCM-AI mở bên trong Excel, không phải nút trên web local. Có thể dùng add-in để tạo snapshot, hoặc tải JSON trên web và đính kèm vào phiên ChatGPT. Sheet nguồn là bản sao tại thời điểm gửi; khi đổi thư viện phải gửi lại. Skill phải báo rõ nếu chưa có nguồn truy cập được, không tạo giá.

## Lưu trữ

Tất cả file giá vẫn ở máy: bản gốc tại `uploads/price-library/<id>/`, thư viện hiện hành tại `libraries/<nghiệp-vụ>/active.json` và các phiên bản tại cùng thư mục. Web chỉ phục vụ loopback, không đưa dữ liệu lên máy chủ công khai. Skill dùng AI của phiên chat để thực hiện yêu cầu; việc nạp giá lưu dữ liệu có phiên bản, không huấn luyện lại trọng số mô hình.
