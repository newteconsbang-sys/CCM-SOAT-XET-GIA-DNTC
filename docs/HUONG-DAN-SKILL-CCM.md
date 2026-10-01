# Skill Soát xét giá CCM

Skill được cài cho Codex với tên `soatxetgia-ccm`. Nếu danh sách chưa hiện, mở lại phiên Codex.

Lệnh ví dụ (thay cột bằng cột bạn chọn):

> Dùng $soatxetgia-ccm, nghiệp vụ ĐNTC. Dò vùng CCM đang chọn. Mô tả ở B, đơn vị D, vật tư ở P, nhân công Q. Điền giá theo thư viện, giữ nguyên dòng Lô.

Nạp thư viện:

> Dùng $soatxetgia-ccm, nạp file thư viện này vào EW. Cột mô tả…, đơn vị…, thương hiệu…, vật tư…, nhân công…

Bốn thư viện riêng ở `D:\00.Setupgpt\SoatXetGia\libraries\DNTC`, `EW`, `NCC`, `NTP`. Mỗi thư mục có mẫu `Mau-thu-vien.csv`. Nhập giá số không có dấu phân cách hàng nghìn. Không có giá thì để trống. CSV mẫu chưa có dữ liệu giá thật.

Skill là hướng dẫn và công cụ quản lý thư viện, dùng AI của phiên chat. Nó không tự tạo bảng giao diện Excel, không tự cài vào khung ChatGPT Excel và không tự huấn luyện AI cục bộ. Để đọc/ghi workbook đang mở cần kết nối Excel được nhận diện thực tế.

Muốn dùng trong khung ChatGPT Excel, có thể dán nội dung SKILL.md làm yêu cầu và cung cấp thư viện qua file hoặc vùng Excel được kết nối. Không giả định chat đó đọc được thư mục ổ D. Để có nút nạp thư viện/chọn cột/dò giá như ảnh ban đầu, cần add-in làm giao diện và kết nối xử lý; bản skill này không thay thế phần đó.

## Giao diện Excel CCM-AI 0.4

Mở lại bảng CCM-AI trong Excel để tải bản mới. Chọn nghiệp vụ. Trong mục Thư viện riêng, chọn vùng giá nguồn trên Excel, lấy vùng và xác nhận cột rồi lưu. Trong Cài đặt cột trên bảng CCM, nhập tên cột hoặc chọn một ô của cột trên Excel và bấm Lấy cột từ ô đang chọn. Chọn vùng CCM có hai cột giá và các dòng cần điền, bấm dò thư viện, kiểm tra nguồn, xem trước rồi Điền giá vào CCM. Dòng Lô giữ nguyên.

Add-in và skill dùng chung bốn thư mục thư viện. Add-in vẫn có AI cục bộ trong mục Soát xét mở rộng; việc điền trực tiếp dùng giá thư viện đã nạp, không tự lấy giá từ kiến thức AI. Bản cập nhật đã kiểm tra biên dịch và dịch vụ; cần kiểm tra đọc/ghi trong Excel thật trước khi dùng trên hồ sơ chính thức.
