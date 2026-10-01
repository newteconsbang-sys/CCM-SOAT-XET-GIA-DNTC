# Cài CCM-AI trên Excel Windows

## Máy đã kiểm tra

Microsoft 365 Apps for enterprise (O365ProPlusRetail), Excel 64-bit, phiên bản 16.0.20430.20118. Có Excel.exe và quyền đọc cấu hình Office. Không phát hiện dữ liệu/dự án cũ trong D:\SoatXetGia; D:\00.Setupgpt\SoatXetGia chưa tồn tại.

## Cài bằng nút bấm

1. Giải nén bộ bàn giao vào `D:\00.Setupgpt\SoatXetGia` hoặc dùng thư mục hiện tại `D:\SoatXetGia`. Không giải nén đè lên dữ liệu cũ. Bộ cài dùng đường dẫn tương đối; không cố định thư mục máy khác.
2. Bấm hai lần **Cai-dat-CCM.cmd**. Python đã đóng gói trong `runtime/python`; không cần Docker, Node hoặc Codex khi sử dụng. Nếu Windows hiện cảnh báo chứng chỉ, kiểm tra tên **CCM-AI localhost**, rồi chọn Có/Yes để tin cậy chứng chỉ HTTPS riêng của dịch vụ cục bộ. Bộ cài chỉ đăng ký add-in CCM và chứng chỉ CCM cho tài khoản hiện tại.
3. Lưu tài liệu đang mở, tự đóng và mở lại Excel. Bộ cài không tự đóng Excel đang dùng.
4. Bấm shortcut **Khởi động CCM-AI** trên Desktop. Nếu chưa tạo được shortcut, bấm **Khởi động CCM-AI.cmd** trong thư mục dự án.
5. Trong Excel: **Home/Trang đầu → Add-ins/Phần bổ trợ → Advanced/Nâng cao → My Add-ins/Phần bổ trợ của tôi → Developer Add-ins → CCM-AI → Add/Thêm**. Tên nhóm có thể thay đổi theo bản Excel. Sau khi thêm, tab **CCM-AI** có nút **CCM-AI – Soát xét giá**, mở bảng điều khiển bên phải.
6. Nếu giao diện ghi **Excel đã kết nối**, lấy vùng thử và thực hiện luồng trong HUONG-DAN.md. Chỉ trạng thái này chưa chứng minh ghi công thức thành công; cần hoàn thành kiểm thử đọc/ghi bên dưới.

## Nếu không thấy Developer Add-ins

Quản trị viên có thể chặn add-in hoặc việc tin cậy chứng chỉ. Không thay đổi chính sách doanh nghiệp. Đưa `manifests/ccm-ai.xml` cho IT để xem phương án triển khai quản trị. Đối với nghiệm thu cá nhân, có thể dùng catalog thư mục chia sẻ:

1. Chia sẻ **chỉ** thư mục `manifests` cho chính tài khoản của bạn, ghi lại đường dẫn UNC. Không chia sẻ `data`, `uploads`, `backups` hoặc toàn bộ dự án.
2. Excel → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs. Thêm UNC vào Catalog URL, chọn Show in Menu.
3. Tự đóng/mở Excel; Home → Add-ins → Advanced → Shared Folder; chọn CCM-AI.

Microsoft mô tả catalog chia sẻ là phương thức kiểm thử, không hỗ trợ triển khai add-in sản xuất bằng cách này. Bản bàn giao này là bản cá nhân để nghiệm thu; triển khai sản xuất cần IT đánh giá phân phối, chính sách chứng chỉ và sao lưu.

Nguồn: [Hướng dẫn sideload Windows của Microsoft](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins), [mã đăng ký sideload Microsoft](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-settings/src/dev-settings-windows.ts), [yêu cầu chạy add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## Kiểm tra kết nối và chứng chỉ

- Shortcut kiểm tra chứng chỉ còn hạn, kho tin cậy, cổng 3443 và `/api/health`; chỉ mở Excel khi dịch vụ đúng tên CCM-AI đã trả lời.
- Dịch vụ chỉ nghe trên `127.0.0.1`, truy cập bằng `https://localhost:3443`. Không mở cổng mạng LAN.
- Nếu cổng đã bị dịch vụ khác dùng, không tự dừng tiến trình khác. Xem `data/server-error.log`.
- Office.js tải từ CDN chính thức Microsoft; cần kết nối đến `appsforoffice.microsoft.com` để Excel khởi tạo Office.js. Hồ sơ không gửi lên CDN này.
- Chứng chỉ có hạn một năm; bộ khởi động chặn khi còn dưới 7 ngày. Khi hết hạn cần lưu lại cặp chứng chỉ cũ, tạo cặp mới và tin cậy lại; không bỏ qua kiểm tra HTTPS.

## Gỡ cài đặt

Chạy `scripts/Uninstall-CCM.ps1` bằng PowerShell (Run with PowerShell). Gỡ đúng đăng ký CCM, shortcut và chứng chỉ ghi nhận trong runtime.json; giữ nguyên dữ liệu, tài liệu, báo cáo và bản sao. Không xóa cache Office chung hay add-in khác.

## Nghiệm thu tích hợp

`tests/CCM-AI-Acceptance.xlsx` là workbook kiểm thử riêng từ mẫu sideload của Microsoft, không phải dữ liệu tài chính thực tế. Khi được Excel nạp cùng CCM-AI, runner chỉ chạy nếu thuộc tính Title đúng `CCM-AI-ACCEPTANCE-ONLY`. Runner tạo các sheet tên TEST, đọc qua Office.js, ghép khác thứ tự, ghi công thức thật, đối soát 1080/120, chặn ghi đè và khôi phục. Kết quả được lưu vào `docs/local-officejs-acceptance.json`. Không đổi Title của workbook thật thành giá trị này.

Nếu file kết quả chưa tồn tại hoặc có kiểm tra failed, chưa được coi tích hợp Excel thành công. Người dùng cũng có thể làm các bước thủ công trong NGHIEM-THU.md.
