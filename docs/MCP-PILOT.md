# Thử kết nối CCM-AI trực tiếp với ChatGPT

Bản thử chỉ công bố công cụ `ccm_pilot_read_one_dntc`. Công cụ đọc đầu mục đầu tiên
của thư viện ĐNTC đang dùng trên máy, kèm nguồn, phiên bản và SHA256. Không sửa
thư viện, không đọc/ghi workbook, không công bố các công cụ đọc toàn bộ thư viện.

## Đã kiểm tra tại máy

Chạy `scripts/test_mcp_pilot.py` bằng Python CCM-AI. Báo cáo ở
`backups/mcp-pilot/local-result.json` xác nhận khởi động stdio, tìm công cụ,
đọc một đầu mục, từ chối gọi thư viện khác/đọc nhiều dòng và giữ nguyên thư viện.
Đây là bằng chứng phần MCP cục bộ; chưa phải bằng chứng ChatGPT Excel đã kết nối.

## Phần tài khoản còn cần thiết lập

1. Tạo Secure MCP Tunnel tại trang Platform, chọn đúng tổ chức và workspace
   ChatGPT của bạn. Tạo hoặc sử dụng runtime API key có quyền Tunnels Read + Use.
   Không dùng admin key để chạy cầu nối, không gửi khóa trong chat.
2. Chạy PowerShell tại dự án với ID thật đã tạo:
   `./scripts/Start-CCM-PilotTunnel.ps1 -TunnelId <ID thật>`.
   Script hỏi khóa bằng ô nhập ẩn, dùng khóa trong bộ nhớ tiến trình và không ghi
   khóa vào cấu hình. Giữ terminal mở trong lúc thử; Ctrl+C dừng cầu nối.
   Cấu hình đã tồn tại sẽ không tự bị ghi đè; khi cần chạy lại, dùng tunnel-client
   run với profile đã tạo và khóa trong biến môi trường của phiên terminal.
3. Khi client đã khỏe và sẵn sàng, thêm kết nối developer-mode ở ChatGPT Plugins,
   chọn Connection → Tunnel và ID đó. Không coi localhost:3443 là URL MCP remote.
4. Chọn kết nối trong phiên ChatGPT và gửi:
   “Thử kết nối CCM-AI: gọi ccm_pilot_read_one_dntc, chỉ trả mô tả, đơn vị, giá
   vật tư, giá nhân công, phiên bản và nguồn. Không đọc sheet thư viện và không
   sửa workbook. Nếu không có công cụ thì báo chưa kết nối.”
5. Chỉ đánh dấu thành công sau khi thấy lượt gọi công cụ và kết quả trong khung
   ChatGPT Excel. Khung Excel có Skills không chứng minh đã hỗ trợ kết nối MCP.

Plugin tài khoản hiện tại vẫn là gói skill. Không tải stdio lên tài khoản rồi
báo đã có kết nối. Sau thử thành công mới gắn dependency kết nối thật vào plugin
hiện tại và mở rộng phạm vi đọc các thư viện theo yêu cầu.

Tham khảo chính thức:
- https://developers.openai.com/api/docs/guides/secure-mcp-tunnels
- https://developers.openai.com/plugins/deploy/connect-chatgpt
- https://github.com/openai/tunnel-client/releases/latest

Tunnel-client v0.0.15 Windows amd64 tải từ nguồn OpenAI, SHA256 ZIP đã đối chiếu:
`3b53133a1e24d43f63088d843860cb1701a4c3ed6390de2e19f69089e43bddc1`.
## Kết quả thực tế ngày 02/10/2026

- Đã tạo Tunnel `tunnel_6abf80d2a6e48191a9302aa41bb6255d` cho tổ chức Personal
  và workspace ChatGPT được tài khoản liệt kê.
- Đã tạo khóa runtime chỉ có Tunnels Read + Use, hết hạn sau 1 ngày. Khóa được
  bảo vệ bằng Windows DPAPI ở máy; không lưu khóa rõ vào file hoặc chat.
- Đã kết nối plugin thử `CCM-AI pilot ĐNTC` vào tài khoản, sau khi người dùng
  xác nhận riêng sử dụng No authentication ở lớp MCP bên trong Tunnel riêng.
- Client được quản lý bằng `tunnel-client runtimes connect`, báo process_running,
  healthy và ready đều true. Báo cáo ở `backups/mcp-pilot/tunnel-status.json`.
- Công cụ thử đã xuất hiện trong phiên chat này. Lượt gọi qua connector trả lỗi
  `MCP SSE probe returned 404 from openai.org` (UNAVAILABLE). Do vậy chưa xác nhận
  lấy giá qua ChatGPT và chưa xác nhận khung Excel sử dụng được cầu nối.
- Không sửa workbook hoặc thư viện. Plugin skill cũ vẫn giữ nguyên.

Khởi động lại client với script `scripts/Connect-CCM-PilotTunnel.ps1` và Tunnel ID
ở trên; script giải mã khóa trong bộ nhớ rồi dùng runtime được quản lý. Không
chạy lại bước tạo khóa hay tạo Tunnel. Khóa thử sẽ hết hạn; chưa cấu hình tự gia
hạn hoặc khởi động cùng Windows.
