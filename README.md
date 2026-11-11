# CCM-AI – Soát xét giá

Web local và Excel add-in cá nhân cho Microsoft 365 trên Windows. Giao diện React/TypeScript, FastAPI/SQLite chạy HTTPS trên loopback; runtime Python được đóng gói riêng. Web quản lý thư viện dùng độc lập; skill trong Codex trên máy đọc trực tiếp giá đã lưu để soát hồ sơ.

**Web local như prepareccm:** mở shortcut Desktop **CCM-AI Soat gia (Web)** hoặc `Mo-web-CCM-AI.cmd`, rồi vào https://localhost:3443/. Chọn ĐNTC/EW/NCC/NTP, nạp file, chọn các cột, xem trước và xác nhận lưu. Xem [hướng dẫn web và skill](docs/WEB-LOCAL-CCM-AI.md).

Bản 0.7.0 tự chọn bảng, nhận diện cột bằng AI Qwen cục bộ và hiển thị **Hạng mục, Đơn vị, Khối lượng, Đơn giá vật tư, Đơn giá nhân công**. Khối lượng thiếu giữ trống; giá và khối lượng lấy từ file gốc, không do AI tạo. Có xem trước tự động, chỉnh cột khi cần và tải Excel đúng 5 cột.

Bản 0.6.1 nhận thêm **PDF có bảng chữ**, chuyển bảng sang XLSX để tải về, gộp bảng cùng tiêu đề qua nhiều trang và lưu giá đã ánh xạ với nguồn trang PDF. Người dùng chọn định dạng dấu chấm/phẩy của giá và đối chiếu bản gốc trước lưu. PDF ảnh scan cần OCR, chưa được đọc tự động.

**Excel add-in:** bấm `Cai-dat-CCM.cmd`, xác nhận chứng chỉ CCM-AI localhost trực tiếp trong Windows, rồi bấm shortcut Khởi động CCM-AI. Xem [hướng dẫn cài đặt](docs/CAI-DAT.md) và [hướng dẫn sử dụng](docs/HUONG-DAN.md).

**Trạng thái nghiệm thu:** bản 0.6 đã kiểm tra luồng web nạp/lưu/đọc thư viện và helper của skill bằng dữ liệu thử tách biệt. Xem [kiểm tra web](docs/KIEM-TRA-WEB-0.6.md) và [kết quả các bản trước](docs/KET-QUA-KIEM-TRA.md). Chưa xác nhận thao tác ghi giá vào workbook thật trong Excel.

Xem [AI cục bộ](docs/AI-CUC-BO.md) để cài mô hình và duyệt gợi ý trong Excel.

## Các phần đã triển khai

- Tab/nút CCM-AI trong manifest; bảng điều khiển theo giao diện tham chiếu, bốn nghiệp vụ.
- Đọc hai vùng, giá trị/công thức/chuỗi hiển thị, danh sách sheet; ánh xạ linh hoạt, cột ĐVT/KL ngoài vùng, mẫu ánh xạ lưu SQLite.
- Ghép theo STT/mô tả/đơn vị với xác nhận khi thiếu hoặc mơ hồ; giữ mã chuỗi, loại dòng nhóm/tổng.
- Giá đơn, giá gộp độc lập VT/NC, tỷ lệ từ ô chọn; không tự đặt hệ số.
- Xem trước, chặn chồng nguồn, dữ liệu đổi, bảo vệ, công thức lỗi/vòng; ghi giá trị/công thức thật; bản sao lưu trên đĩa và khôi phục.
- Tài liệu XLSX/DOCX/PDF theo ba khu vực, lưu nguyên gốc/hash/phiên bản và trích chữ; nguồn giá có metadata và điều kiện so sánh.
- Quy tắc/mẫu riêng theo nghiệp vụ có phiên bản; biểu mẫu chỉnh bằng nút bấm; xuất XLSX/DOCX/PDF và JSON căn cứ.
- Shortcut, kiểm tra HTTPS/dịch vụ, script cài/gỡ; Python portable; Git bỏ qua dữ liệu nhạy cảm.

## Giới hạn cụ thể

Đã có AI cục bộ qua Ollama để gợi ý ghép ngữ nghĩa và diễn giải có căn cứ; xem docs/AI-CUC-BO.md. Chưa có OCR, xác minh kỹ thuật tự động, trình dựng mọi công thức Excel, hộp Save As Windows riêng, sao lưu toàn bộ thư viện tự động. Công thức kiểm tra hỗ trợ tham chiếu A1 với số học và SUM/AVERAGE/MIN/MAX/COUNT/IF/ROUND/ABS; tên vùng, tham chiếu ngoài workbook, hàm động bị chặn. Cách tính thành tiền khác KL×(VT+NC) cần giữ nguyên file gốc và chưa được tính tự động. Quy tắc nghiệp vụ thực tế chưa được cung cấp nên khởi tạo chưa duyệt, cần người phụ trách hiệu chỉnh.

## Mã nguồn và build

`frontend/src`: UI, xử lý dòng, Office.js và runner nghiệm thu. `backend`: API, SQLite, chứng chỉ và xuất báo cáo. `manifests`: XML add-in. `scripts`: cài/khởi động/gỡ/đóng gói. `tests`: kiểm thử và workbook sideload riêng. Các thư mục dữ liệu theo yêu cầu được tạo tự động, không đưa vào Git.

Dev dùng Node 22+, pnpm và Python 3.12. `frontend/package.json` và lockfile xác định thư viện frontend; `backend/requirements.txt` xác định thư viện backend. Bản bàn giao đã có `frontend/dist`, `backend/vendor` và `runtime/python` nên người dùng không chạy build.

Khi phát triển, chạy tsc và Vite trong frontend, chạy `tests/core.test.ts` với tsx, chạy `tests/backend_test.py` với runtime Python. Backend kiểm thử cần httpx2 (chỉ phục vụ test). Không commit `config`, khóa/chứng chỉ riêng, SQLite, uploads, reports hay backups. Không cấu hình remote Git hoặc gửi mã/dữ liệu ra mạng trong bộ bàn giao.


Bản 0.3 bổ sung Ghi nhớ nghiệp vụ đã duyệt: lưu ví dụ chuẩn trong SQLite, tra cứu theo nghiệp vụ/ĐVT/mô tả và hiển thị bài học cho lần AI tiếp theo; có nút thu hồi. Chưa huấn luyện lại trọng số mô hình. Xem docs/AI-CUC-BO.md.

Bản 0.5: nút **Gửi thư viện cho skill ChatGPT** đưa thư viện được chọn vào sheet workbook có phiên bản và nguồn; xem [cách kết nối](docs/KET-NOI-CHATGPT-EXCEL.md). Dịch vụ khởi động khi đăng nhập Windows, không mở Excel/Ollama. Bốn nghiệp vụ cần nạp thư viện thực tế trước khi dùng.


## Bản 0.9 – Thư viện giá & Soát xét giá (web local)

**Thư viện giá** (`#/library`): chọn file Excel → tool tự nhận diện cột, kiểm tra (thiếu tên/ĐVT/giá, giá âm, trùng mã, VT+NC+Máy lệch Đơn giá tổng), lưu thành version mới và áp dụng ngay. Quản lý Vật tư / Nhân công / Ca máy (cột Loại, hoặc tự đoán theo tên sheet, mã NC-/CM-, ĐVT công/ca). Sửa trực tiếp từng ô (tự tính lại Đơn giá tổng, ghi lịch sử người sửa), thêm/xóa mục, lọc theo loại/nhóm, tìm không dấu. So sánh hai version: thêm mới / thay đổi (chỉ rõ trường) / bị xóa, xuất Excel. Xuất version đang xem (theo bộ lọc).

**Soát xét giá** (`#/package/<id>`): đọc BOQ/báo giá NTP, nhận tiêu đề 2 tầng có ô gộp (Đơn giá → Vật tư/Nhân công/Khấu hao), số dạng 1.250.000. Lưới 2 phần: Phần 1 giá NTP báo (STT, Hạng mục, ĐVT, Khối lượng, Đơn giá VT, NC, Khấu hao, Đơn giá tổng, Thành tiền, Ghi chú) – Phần 2 giá KSCP–MEP soát xét (sửa trực tiếp KL, VT, NC, Khấu hao, Ghi chú; tự tính Đơn giá tổng, Thành tiền, Chênh lệch). Đối chiếu version thư viện của gói; NTP cao hơn thư viện quá ngưỡng thì đề xuất hạ về giá thư viện. Kết quả: Hợp lý (xanh) / Giảm giá (đỏ) / Điều chỉnh (vàng). Lưu chỉnh sửa vào SQLite kèm lịch sử; soát xét lại hoặc nạp file mới vẫn giữ các dòng đã sửa; ô đang sửa được giữ nháp trên trình duyệt.

**Xuất Excel**: mở bản gốc lúc nạp, giữ nguyên sheet, công thức, định dạng, ô gộp, dữ liệu; chỉ thêm cột "KSCP–MEP SOÁT XÉT" bên phải sheet báo giá (KL, VT, NC, Khấu hao, Đơn giá tổng, Thành tiền, Chênh lệch dạng công thức; Kết quả tô màu; Ghi chú) và dòng tổng. Giới hạn của openpyxl: hình ảnh/biểu đồ nhúng trong file gốc có thể không được giữ – tool cảnh báo khi file gốc có hình. "Bảng so sánh" xuất bảng gọn riêng.

Module: `backend/library_service.py` + `library_router.py` (thư viện), `boq_reader.py` (nhận diện BOQ), `review_service.py` (tính toán, lưu, lịch sử), `review_export.py` (xuất vào file gốc), `xlsx_utils.py` (đọc số, ghi Excel). Kiểm thử: `tests/test_library_review.py` chạy trên database tạm. Chạy hoàn toàn local, không dùng Ollama hay dịch vụ AI ngoài.

**0.9.1 – Giữ nguyên cấu trúc file:** File thư viện hoặc báo giá NTP có nhiều sheet thì mỗi sheet là một tab, hiển thị đúng cột gốc (tiêu đề, thứ tự), đúng thứ tự dòng kể cả dòng nhóm/tổng (I, II, Cộng…). Soát xét giá tự nhận mọi sheet có bảng giá (bỏ qua Bìa, Tổng hợp không có đơn giá); tab Cấu hình cột cho thêm/bỏ từng sheet. Xuất thư viện (không lọc) trả lại đúng file gốc với các ô đã sửa; xuất soát xét thêm cột soát xét vào từng sheet đã soát xét. Script khởi động tự khởi động lại dịch vụ khi phiên bản mã nguồn (backend/VERSION) mới hơn bản đang chạy; có thêm Khoi-dong-lai-CCM-AI.cmd.

**0.9.5 – % khấu hao vật tư:** Cột kiểu "% KH VẬT TƯ" trong báo giá NTP được nhận diện là tỷ lệ khấu hao (0,2 = 20%) và hiển thị dạng %. Đơn giá = VT × %KH + NC + Khấu hao; Phần 2 có cột "% KH ✎" sửa được (nhập 20% hoặc 0,2), Excel xuất thêm cột "% KH VT soát xét" với công thức tương ứng. Gói đã soát xét trước đó: bấm Tải dữ liệu để nhận cột %KH.

**0.9.6 – Thanh công thức:** Bấm vào ô Đơn giá, Thành tiền, Chênh lệch (cả phần NTP) hiện công thức và số liệu trên thanh fx, ví dụ "= ĐG VT × %KH + ĐG NC = 14.950 × 20% + 0 = 2.990", "= Đơn giá × Khối lượng". Excel xuất dùng công thức rõ ràng: Đơn giá = VT*%KH+NC, Thành tiền = Đơn giá*Khối lượng.

**0.9.9 – Cấu hình cột 2 bên, cột tự tạo, copy/paste:** Tab Cấu hình cột chia 2 bên: Cấu hình báo giá NTP (xếp theo thứ tự cột file) và Cấu hình P.KSCP soát xét (thứ tự cột Phần 2). Nút "+ Tạo cột" → Lưu cột: thêm cột soát xét (Số/tiền hoặc Chữ) vào tab Chi tiết, nhập/sửa được, có lịch sử và xuất Excel. Tab Chi tiết: kéo chuột hoặc Shift+bấm để chọn vùng ô, Ctrl+C copy, Ctrl+V dán (dán 1 giá trị vào cả vùng), Delete xóa vùng; không còn kéo-thả làm di chuyển dữ liệu. Bỏ nút Bảng so sánh.

**0.9.10:** Phần KSCP soát xét gồm các cột: KL, ĐG VT, ĐG NC, % KH, ĐG tổng, Thành tiền, Chênh lệch, Giá TV, Kết quả, Ghi chú (bỏ cột Khấu hao; % KH luôn hiện, để trống = 100%). Excel xuất cùng thứ tự.

**0.9.11 – Kéo thả thứ tự cột, công thức kiểu Excel:** Cấu hình P.KSCP soát xét: kéo ⋮⋮ (hoặc ↑ ↓) đổi thứ tự mọi cột Phần 2, áp dụng cho tab Chi tiết và Excel xuất. "+ Tạo cột" có 3 kiểu: Số/tiền, Chữ, Công thức. Công thức dùng tên cột trong ngoặc vuông ([KL], [ĐG VT], [ĐG NC], [%KH], [ĐG tổng], [Thành tiền], [Chênh lệch], [Giá TV], [KL NTP], [ĐG VT NTP], [ĐG NC NTP], [ĐG NTP], [TT NTP], tên cột tự tạo) và hàm SUM, SUBTOTAL, ROUND, ROUNDUP, ROUNDDOWN, IF, AND, OR, MIN, MAX, AVERAGE, ABS; dấu ; tách đối số khi dùng dấu phẩy thập phân. Dòng nhóm / dòng tổng hiện tổng (SUBTOTAL) các cột số và công thức. Ô vàng gõ = để tính (=15.000*1,1 hoặc =[ĐG VT NTP]*0,9). Chọn vùng ô: thanh dưới hiện Tổng / TB / Đếm / Min / Max. Excel xuất cột công thức dạng công thức thật. Chặn hẳn kéo-thả chữ làm mất số trong ô.

**0.9.12 – Thêm / chèn dòng, chữ cột A B C:** Tab Chi tiết có nút "+ Thêm dòng" (dưới dòng đang chọn), "Chèn dòng" (trên), "Xóa dòng" (chỉ dòng KSCP thêm), chuột phải vào dòng cũng có menu. Dòng thêm (dấu + màu xanh) nhập được STT, Hạng mục, ĐVT và các cột soát xét; được giữ khi soát xét lại; Excel xuất ghi thành khối "CÁC DÒNG KSCP BỔ SUNG" dưới bảng (không chèn giữa để giữ công thức file gốc). Hàng chữ cột A, B, C… như Excel (Phần 1 đúng cột file gốc, Phần 2 đúng cột trong file xuất). Ô chỉ sáng lên khi đang thao tác, sửa xong trở lại màu bình thường.

**0.9.13 – Hàm Excel ngay trên bảng Chi tiết:** Gõ = vào ô vàng như Excel, dùng địa chỉ ô theo hàng chữ cột (A B C…) và cột Dòng: =M7*1,1 · =SUM(E7:E12) · =SUBTOTAL(9;Q7:Q20) · =ROUND(M7*%;0) · =$M$7 · =[ĐG VT NTP]*0,9. Ô có công thức có dấu tam giác tím ở góc; bấm vào ô hiện lại công thức để sửa, thanh fx hiện công thức → kết quả; ô tự tính lại khi ô nguồn đổi; công thức được lưu cùng gói. Gõ số thường đè lên thì bỏ công thức.

**0.9.14:** Bỏ các nút Thêm/Chèn/Xóa dòng trên thanh công cụ – thao tác bằng chuột phải vào dòng.

**0.9.15 – Gõ hàm như Excel:** Thanh fx sửa được: bấm ô vàng, gõ số hoặc công thức trên thanh fx, Enter để ghi. Đang gõ công thức (sau dấu = + - * / ( ; ,) bấm vào ô khác trên bảng → tự chèn địa chỉ ô (M7), kéo chuột → chèn vùng (E7:E9), giống Excel.

**0.9.16:** Cột Giá TV đổi thành % Chênh lệch = Chênh lệch / Thành tiền NTP, có ▲ tăng (cam) / ▼ giảm (đỏ); dòng nhóm, dòng tổng cũng có. Excel xuất thêm cột "% chênh lệch" dạng công thức với định dạng ▲/▼. Giá thư viện vẫn xem khi rê chuột lên Hạng mục, nút Áp giá thư viện giữ nguyên.

**0.9.17:** Kết quả đổi thành: NTP cao hơn soát xét → "Cao" (đỏ); NTP thấp hơn → "Thấp" (xanh dương); bằng → "Hợp lý" (xanh lá). % Chênh lệch: ▼ đỏ khi NTP cao, ▲ xanh dương khi NTP thấp. Excel xuất cùng màu.

**0.9.18 – Hàm Excel, gợi ý hàm, làm tròn đơn giá:** Bảng Chi tiết hỗ trợ + - * / ^, SUM, SUMPRODUCT, SUBTOTAL (9 tổng, 1 TB, 2 đếm, 4 max, 5 min, 6 tích), ROUND/ROUNDUP/ROUNDDOWN, IF, MIN, MAX, AVERAGE, COUNT, PRODUCT, ABS, AND, OR. Gõ = rồi chữ cái → hiện gợi ý hàm (↑ ↓, Tab/Enter chèn) và cú pháp hàm đang gõ. Đơn giá soát xét mặc định làm tròn đến trăm đồng (…00): gõ số lẻ hay công thức đều tự làm tròn; chỉnh ở Cấu hình cột (trăm / nghìn / chục / không). Excel xuất: ĐG tổng =ROUND(VT*%KH+NC,-2), ô gõ công thức xuất thành công thức thật (đơn giá bọc ROUND), địa chỉ ô trên web trùng với file xuất. Phần 2 chọn ô / vùng ô như Phần 1.

**0.9.19 – Sửa số NTP bằng công thức, lịch sử theo lần Lưu:** Các ô số Phần 1 (Khối lượng, ĐG VT, ĐG NC, %KH, Đơn giá) sửa được và gõ công thức như Excel (=E7*0,1, =SUM(…)); sửa VT/NC/%KH thì Đơn giá NTP tự tính lại; ô đã sửa chữ tím nghiêng, rê chuột thấy số NTP gốc; Excel xuất ghi số/công thức vào đúng ô gốc (giữ công thức Đơn giá của file). ĐG tổng soát xét gõ thẳng được (bỏ tách VT/NC). Lịch sử chỉnh sửa chỉ ghi khi bấm Lưu; mỗi lần Lưu tạo một bản (số liệu + file Excel lưu tại data\review_versions): tab Lịch sử có danh sách Các lần lưu → Xem (mở lại bảng lúc đó, chỉ xem) / Tải file.

**0.9.20 – Chọn cả dòng / cả cột như Excel:** Bấm số ở cột Dòng → chọn cả dòng (kéo hoặc Shift+bấm để chọn nhiều dòng); bấm chữ cột A B C… → chọn cả cột (kéo / Shift để chọn nhiều cột); copy, xóa, xem Tổng/TB/Đếm như vùng chọn thường. Chuột phải vào chữ cột phần soát xét → Chèn cột bên trái / phải (Số hoặc Chữ), Xóa cột tự tạo.

**0.9.21:** File xuất: cột ngăn cách phần NTP và phần soát xét tô xanh nhạt, rộng 2.

**0.9.22 – Thao tác ô như Excel:** Bấm ô nào thì ô hiện giá trị, thanh fx hiện công thức của đúng ô đó (bấm ô không sửa được cũng hiện giá trị ô đó); gõ là thay nội dung, F2 hoặc nhấp đúp để sửa công thức ngay trong ô; rời ô không đổi gì thì giữ nguyên công thức. Ctrl+C / Ctrl+V chép cả công thức và tự dịch địa chỉ (=E7*0,1 dán xuống dòng dưới thành =E8*0,1; $E$7 giữ nguyên). Kéo ô vuông nhỏ góc dưới-phải ô / vùng chọn xuống để chép công thức hàng loạt; Ctrl+D điền xuống trong vùng chọn.

**0.9.23 – Thanh công cụ kiểu Excel trên bảng Chi tiết:** Hoàn tác / Làm lại (Ctrl+Z / Ctrl+Y), Copy, Dán, Điền xuống; Đậm / Nghiêng / Gạch chân (Ctrl+B/I/U), màu chữ, màu nền (bảng màu + 4 màu nhanh), căn trái / giữa / phải, định dạng số (1.234 / 1.234,00 / %), Xóa định dạng – áp cho ô / vùng / cả dòng / cả cột đang chọn, lưu cùng gói (bấm Lưu) và xuất ra Excel đúng ô.

**0.9.24 – Tăng tốc:** Lưu nhanh hơn ~8 lần (file Excel của bản lưu xuất nền), Xuất Excel nhanh hơn ~10 lần (gán kiểu ô qua chỉ số cache); bảng Chi tiết chỉ vẽ lại dòng có thay đổi (chọn ô, kéo chọn, bấm ô không vẽ lại cả 400+ dòng), bỏ position cho mọi ô (giảm thời gian vẽ).

**0.9.25:** Đang gõ công thức (=SUM( …) và bấm / kéo chọn ô: các ô được tham chiếu tô viền màu + chữ đậm như Excel (mỗi vùng một màu), chỉ tô – không sửa giá trị ô nào; Enter xong thì bỏ tô. Thanh fx hiện đúng nội dung đang gõ trong ô.
Thanh fx có nhãn cho ô đang chọn: "ƒ Công thức" (tím), "123 Giá trị" (xám) hoặc "⚙ Tự tính" (xanh – ĐG tổng, Thành tiền, Chênh lệch…).

**0.9.26:** Ô tool tự tính hiện đúng công thức Excel (giống file xuất) và nhãn "ƒ Công thức": ĐG tổng =ROUND(M14*O14+N14;-2), Thành tiền =P14*L14, Chênh lệch =Q14-I14, % chênh lệch =R14/I14. Bấm Enter trên thanh fx mà không sửa thì giữ nguyên.

**0.9.27:** Chuột phải vào ô: tự chọn ô đó (nếu ngoài vùng chọn) và menu có Copy (Ctrl+C), Dán (Ctrl+V, kèm công thức), Dán giá trị, Xóa nội dung, Chèn dòng trên / dưới. Ctrl+C copy được cả 1 ô.

**0.9.28:** Bấm chọn ô công thức → trong ô hiện và bôi chọn CÔNG THỨC (chữ tím, vd =E8*2, ĐG tổng =ROUND(M14*O14+N14;-2)); ô giá trị → hiện GIÁ TRỊ. Rời ô không sửa thì hiện lại kết quả, giữ nguyên công thức.

**0.9.29:** Nút "Tải dữ liệu" = về hồ sơ gốc đã tải lên: hỏi xác nhận rồi bỏ hết số / công thức / định dạng đã sửa và các dòng đã thêm, đọc lại file NTP và đối chiếu thư viện như lúc mới tạo. Các lần Lưu trước vẫn xem / tải lại ở tab Lịch sử chỉnh sửa. (Đổi cấu hình cột vẫn giữ chỉnh sửa.)

**0.9.30:** Áp giá thư viện / Giữ giá NTP chỉ áp đúng ô đang chọn (chọn cột ĐG VT thì chỉ đổi ĐG VT; chọn cả dòng hoặc tick ô vuông thì áp cả dòng). Giữ Ctrl + bấm / kéo để chọn nhiều vùng rồi áp giá, định dạng, xóa 1 lượt. Copy nguyên dòng (bấm số Dòng → Ctrl+C) rồi chuột phải → "Chèn dòng đã copy (xuống dưới)" hoặc phím Insert: mặc định chèn dưới dòng đang chọn. Ctrl+C / Ctrl+V / Delete / Insert dùng được cả khi chọn bằng số Dòng / chữ cột.

**0.9.31 – Y chang Excel:** Bấm ô: ô vẫn hiện giá trị, công thức chỉ hiện trên thanh fx (F2 / nhấp đúp để sửa trong ô). Ô của file NTP hiện đúng công thức gốc trong file: dòng tổng/nhóm =SUBTOTAL(9,I6:I26), Đơn giá =ROUND(E7*G7+F7,2), Thành tiền =ROUND(H7*$D7,2)… (gói cũ tự đọc bổ sung công thức khi mở).

**0.9.32:** Excel xuất: đơn giá KSCP soát xét mặc định là công thức ROUND(…,-2) – giá giữ theo NTP thì tham chiếu ô NTP (=ROUND(E7,-2)), giá thư viện / nhập tay thì =ROUND(15000,-2); ĐG tổng =ROUND(M7*O7+N7,-2).

**0.9.33:** KL soát xét làm tròn số nguyên: trên bảng tự làm tròn (gõ 207,58 → 208), Excel xuất =ROUND(D7,0) khi theo NTP, =ROUND(số,0) khi sửa (chọn "Không làm tròn" ở Cấu hình cột thì giữ số lẻ).

**0.9.34:** Gõ số / công thức vào ĐG tổng (cột P) chỉ đổi ô đó – ĐG VT, ĐG NC giữ nguyên như Excel; xóa trắng ô ĐG tổng → về công thức mặc định ROUND(VT×%KH+NC,-2). Excel xuất ghi đúng công thức / số của ô ĐG tổng đã nhập.

**0.9.35:** Như Excel: ô đang chọn viền xanh đậm; chọn ô công thức hoặc đang gõ công thức → mọi ô trong công thức tô viền + nền + chữ theo màu riêng từng tham chiếu, lần nào cũng hiện (kể cả ô đã sửa / dòng đang chọn); chọn ô giá trị thì bỏ tô.

**0.9.36:** Bỏ thêm / xóa cột trực tiếp trên bảng Chi tiết (chuột phải chữ cột không còn menu); chữ cột chỉ để chọn cả cột.

**0.9.37:** Bỏ dòng chọn làm tròn ở Cấu hình cột – luôn làm tròn đơn giá đến trăm đồng, KL số nguyên.

**0.9.38:** Excel xuất: tổng soát xét (Thành tiền, Chênh lệch, % chênh lệch) ghi cùng hàng với dòng "TỔNG …" của file NTP và tô xanh nhạt; file không có dòng tổng thì ghi ở hàng dưới cùng như trước.

**0.9.39:** Bỏ tô vàng ô sửa được và tiêu đề ✎ ở phần soát xét – cùng màu với phần NTP; ô đang chọn chỉ viền xanh.

**0.9.40:** Như Excel: công thức gõ bắt đầu bằng = hoặc + (vd +M8*2, + rồi bấm ô M9 rồi +1000), hoặc - kèm địa chỉ ô (-E7*2); tool lưu thành =…, có gợi ý hàm, chọn ô bằng chuột và tô màu ô tham chiếu như khi gõ =.

**0.9.41:** Bỏ dấu tam giác tím ở góc ô công thức (xem loại ô trên thanh fx: ƒ Công thức / 123 Giá trị).

**0.9.42:** Phần soát xét có dòng đậm / nhạt xen kẽ giống phần NTP bên trái.

**0.9.43:** Excel xuất – khối KSCP soát xét: phông Times New Roman, tiêu đề CHỮ HOA cỡ 13, nội dung cỡ 11 (giữ đậm / màu Cao–Thấp–Hợp lý).

**0.9.44:** Excel xuất: dòng tổng (cả phần NTP và phần soát xét) và dòng tiêu đề của cả 2 phần dùng Times New Roman cỡ 13.

**0.9.45:** Excel xuất: dòng hạng mục mẹ (A, 1, 1.1, 1.1.1… và hạng mục có STT con như 1.2.1) ở phần soát xét giữ đúng màu nền của file NTP; dòng mẹ không có màu trong file thì tô xám nhạt. Dòng con giữ định dạng ô của file.

**0.9.46:** Excel xuất: dòng STT mẹ (1, 1.1, 1.1.1, 1.2.1…) ở phần soát xét có Thành tiền / Chênh lệch = SUBTOTAL(9, các dòng con), % chênh lệch theo thành tiền NTP của dòng đó; dòng tổng dùng SUBTOTAL(9,…). Bảng Chi tiết: tiêu đề cột in HOA.

**0.9.47:** Excel xuất tự căn: độ rộng cột soát xét theo nội dung / tiêu đề, trang in khổ ngang vừa 1 trang chiều ngang (A4, bảng rộng thì A3), lề hẹp, căn giữa, lặp dòng tiêu đề mỗi trang.

## 0.9.48
- Bang Chi tiet: keo mep phai chu cot (A, B, C...) de doi do rong cot, keo mep duoi so dong de doi chieu cao dong; nhap dup mep cot de tra ve do rong mac dinh. Kich thuoc duoc nho theo goi + sheet tren may nay.

## 0.9.49
- Bang Chi tiet: dai tieu de PHAN 1 va PHAN 2 to mau do nhat, chu to 1.5 lan.

## 0.9.50
- Tab Lich su chinh sua: bo bang Chi tiet chinh sua, chi giu danh sach Cac lan luu.

## 0.9.51
- Thong tin goi gon lai 2 dong: Ngay tao / Du an / Goi thau / Trang thai; Gia tri NTP bao / Tong gia soat xet / Chenh lech. Bo dong File bao gia NTP (ten file hien khi re chuot vao Goi thau).

## 0.9.52
- Thanh dinh dang: Mau chu va To mau o chon tu bang 7 mau co ban (nhu Excel), co Tu dong / Khong to mau; to mau phu kin toan o.

## 0.9.53
- Bang mau chu / to o: Do, Xanh la cay, Xanh duong, Vang, Luc lam, Hong, Trang, Den, Xam.

## 0.9.54
- Xuat Excel: ke khung toan bo bang (Phan 1, cot ngan cach, Phan 2) tu dong tieu de toi dong cuoi bang net lien mau xam.

## 0.9.55
- Bang Chi tiet: dong tieu de cot in dam.

## 0.9.56
- Copy o co cong thuc san (vd DG tong =ROUND(M7*O7+N7;-2)) se copy cong thuc nhu Excel, dan sang o khac tu dich dia chi; dan lai dung cong thuc san thi bo so go de.

## 0.9.57
- Giu gia NTP / Ap gia thu vien: ap dung dung o dang chon (ke ca o DG tong rieng le, bo cong thuc go tay o do); khong co gia thu vien thi dien 0.

## 0.9.58
- Thu vien gia: nut xoa (thung rac) tren tung version va nut Xoa thu vien; xoa duoc ca ban dang ap dung (ban con lai moi nhat tu thanh hien hanh), goi thau dang dung chuyen sang ban con lai.

## 0.9.59
- Chon o nhu Excel: bam 1 lan chi chon ca o (khong hien con tro / boi chu); nhap dup hoac F2 moi vao sua, o cong thuc hien cong thuc, o gia tri hien gia tri; Esc huy sua; mui ten trai/phai, Tab sang o ben canh; go la thay noi dung; Ctrl+C copy cong thuc neu o la ham, gia tri neu o la so.

## 0.9.60
- Copy o co cong thuc san trong file NTP (vd don gia =ROUND(E7*G7+F7,2)) cung copy cong thuc; Copy/Dan bang nut tren thanh cong cu va chuot phai deu dan cong thuc.

## 0.9.61
- Danh sach goi thau: hien day du so tien (dong) cho Gia NTP bao, Gia soat xet; them cot Chenh lech (dong) ben canh Chenh lech %.

## 0.9.62
- Thu vien gia: chon Phan loai truoc khi nap file (Don gia DNTC chuan / Bao gia NCC / Hop dong tham khao). DNTC chuan ap dung ngay; NCC va Hop dong chi luu. Loc danh sach version theo phan loai, doi phan loai cua version dang xem.

## 0.9.63
- Nap thu vien: nhan tieu de 2 tang (DON GIA (VND) gop o, duoi la VAT TU / NHAN CONG) va cot "DV" la don vi, "CHUNG LOAI" la quy cach - file Don gia DNTC chuan khong con bao "Khong co dong hop le".

## 0.9.64
- Thu vien gia: bo nut Tai file mau va So sanh version tren dau trang.

## 0.9.65
- Thu vien gia: moi sheet trong file nap la 1 tab, giu nguyen cau truc (ca sheet khong co don gia nhu Chi tiet truc, KH vat tu); bo loc Moi loai; khong hien dong loi / sheet bi bo qua, dong loi hien binh thuong.

## 0.9.66
- Thu vien gia: bo cot Loai va So sanh trong bang theo sheet.

## 0.9.67
- Thu vien gia: bo nut Them muc, an dong Ghi chu cuoi sheet don gia; keo mep tieu de de doi do rong cot (nhap dup tra ve tu dong); them phan loai Khau hao, Don gia nhan cong truc.

## 0.9.68
- Thu vien gia: bo o Tat ca nhom, o doi phan loai, nut Lich su sua va Xoa thu vien tren thanh cong cu (xoa bang thung rac o danh sach version).

## 0.9.69
- Thu vien: nap file Khau hao (KH vat tu T=1..T=42) va Don gia nhan cong truc giu nguyen form, la bang tra cuu (khong can cot don gia, khong dat lam hien hanh); hien dong tieu de tren bang (Hang muc, Phuong an) va dong ghi chu dai trai het be ngang, xuong dong de doc.

## 0.9.70
- Thu vien gia: bo cot danh sach ben trai; cac phan loai thanh tab nam ngang (chon tab thi Phan loai nap file chon theo), bang hien full be ngang, chon ban trong tab; them tab File da luu de Xem / Dat hien hanh / Xoa tung file.

## 0.9.71
- Bang thu vien gia: ke o tat ca cac o, tieu de nen xanh la chu dam in hoa, dong nhom to mau theo cap (A / I / I.1), dong xen ke, dong ghi chu nen vang nhat, so canh phai.

## 0.9.72
- Bang thu vien gia: canh giua cot STT.

## 0.9.73
- Bang thu vien gia: canh giua them cot DV va KL.

## 0.9.74
- Khau hao: tool hoc ten vat tu tu ho so BOQ co cot % KH VAT TU (vat tu co the khau hao duoi 100% / mac dinh 100%), da hoc san tu ho so PHU DONG (59 / 52 ten); vat tu chua hoc thi suy theo nhom bang KH VAT TU. Tab Khau hao co bang "Vat tu khau hao da hoc", nut Hoc tu ho so BOQ, chuyen loai tung ten.

## 0.9.75
- File NTP: nhan dien cot Thoi gian bat dau / ket thuc, hien dd/mm/yyyy; tab Cau hinh cot bao thieu cot so voi form chuan.
- Thanh tien NTP (Phan 1) va Thanh tien soat xet (Phan 2) sua duoc / go cong thuc nhu Excel (xoa trang = ve cong thuc KL x Don gia); xuat Excel ghi dung cong thuc.

## 0.9.76
- Thu vien gia: bang co khung cuon rieng, dong tieu de cot dung yen khi keo len xuong.

## 0.9.77
- Sua loi goi da luu cau hinh cot tu file cu (Don gia / Thanh tien tro vao cot Thoi gian): mo goi tu nhan dien lai theo file hien tai, giu cac dong da sua; cot Thoi gian hien dd/mm/yyyy.

## 0.9.78
- %KH soat xet tinh theo thoi gian: so thang = ROUNDUP((Ket thuc - Bat dau)/30,0), tra bang KH VAT TU (thu vien Khau hao moi nhat) theo vi tri du an; vat tu mac dinh 100% -> 100%; o %KH da sua tay giu nguyen.
- Vi tri du an XA BIEN / GAN BIEN: chon khi tao du an, doi o o Vi tri tren dau goi thau (tinh lai %KH).

## 0.9.79
- Thay nut Ap gia thu vien / Giu gia NTP bang o chon nguon + nut Chay: Mac dinh (DNTC chuan, Bao gia NCC nap sau thi ghi de, thieu thi Hop dong tham khao), Don gia DNTC chuan, Hop dong tham khao, Bao gia NCC, Khau hao (%KH theo thoi gian), Don gia nhan cong truc, Giu gia NTP. Ap cho o dang chon, chua chon thi ca bang.

## 0.9.80
- Doi chu nut Chay thanh Run.

## 0.9.81
- Sua loi bam vao o khong chon duoc: bam vao phan trong / vien cua o (dong cao, khoang dem) van chon dung o do.

## 0.9.82
- O thong tin goi thau 2 dong canh thang cot (4 cap nhan / o deu nhau).
- Trang thai goi con 3: Dang soat xet / Da phan hoi BCHCT / Xong (trang thai cu tu doi: Cho duyet, Can giam -> Dang soat xet; Da duyet -> Xong).

## 0.9.83
- O thong tin goi thau: nhan canh trai.

## 0.9.84
- Chon o: moi o (ke ca o chu / o chi xem) bam vao deu co khung xanh, sang so dong + chu cot.
- Dang sua o ma cong thuc / chu dai hon o: o sua tran sang phai nhu Excel.
- Delete / Backspace xoa ca vung chon, ke ca nhieu vung chon bang Ctrl.

## 0.9.85
- Chon o Phan 2: khung xanh ro quanh o dang chon, khong con vet boi chu xam.

## 0.9.86
- Cot % DA KH (file NTP): %KH soat xet = %KH theo bang - %DA KH; dong co % DA KH thi DG NC Phan 2 bi khoa (khong tinh nhan cong).
- Sua truc tiep ngay Bat dau / Ket thuc tren bang (dd/mm/yyyy); bam Luu thi %KH dong do tinh lai; xuat Excel ghi ngay da sua.

## 0.9.87
- Sua truc tiep STT va Mo ta cong viec cua dong trong file (giu ban goc de doi chieu, xuat Excel ghi noi dung da sua).
- Co dinh them cot Don vi (cung STT, Mo ta) khi keo ngang.

## 0.9.88
- Them o He so (canh Vi tri) cho moi goi; Run Don gia nhan cong truc nhan don gia voi he so (vd 2 thap x1,8). Bam Luu de ghi he so.

## 0.9.89
- Sua loi file co cot % DA KH truoc cot % KH VAT TU: tool lay nham % DA KH lam %KH; mo goi tu nhan dien lai.
- %KH soat xet la cong thuc =%KH bang - o % DA KH (vd =0,53-G7 -> 23%).
- O Vi tri | He so chia doi.

## 0.9.90
- Chon Bao gia NCC / Hop dong tham khao: hien o chon file (ten file da nap o Thu vien gia) de Run theo dung file.
- Bao gia NCC: don gia = gia NCC x He so cua goi.

## 0.9.91
- Bao gia NCC dang song ngu (Specifications / Dac tinh ky thuat, Unit / D.vi, Unit price / Don gia): nhan dien dau muc, don vi, don gia lam thu vien gia NCC; luu lai va dung cho Run Mac dinh / Bao gia NCC.

## 0.9.92
- Nap bao gia: chi lay toi dong "Cong truoc thue / Sub-total"; bo dong VAT, tong sau thue, dieu kien thanh toan, chu ky phia duoi.

## 0.9.93
- Bao gia NCC da nap truoc day: tu dong bo cac dong VAT, cong sau thue, dieu khoan, chu ky (chi giu toi dong Cong truoc thue)

## 0.9.94
- Phuong an goi (PA): PA1 Du toan / PA2 BCHCT soat xet / PA3 P.KSCP soat xet - chon khi tao goi va o dau bang
- PA1: file bat buoc du cot A-M; tu ap gia thu vien vao E, F va %KH cot H (tinh tu I, J); Run ghi vao ca Phan 1

## 0.9.95
- Tieu de khoi theo PA: PA1 DU TOAN DNTC (chi Phan 1) / PA2 BAO GIA NTP + GIA BCHCT SOAT XET / PA3 them PHAN 3: P.KSCP SOAT XET
- PA3: Phan 3 lap lai cac cot Phan 2 (KL, DG VT, DG NC, %KH, DG tong, Thanh tien, Chenh lech, %, Ket qua, Ghi chu); o chua sua = theo Phan 2; Run ghi vao Phan 3; xuat Excel them cot Y-AH co cong thuc
- PA1: xuat Excel chi Phan 1, cot Don gia / Thanh tien la cong thuc
- Sua so lon trong cong thuc xuat Excel (khong con dang 1e+06)

## 0.9.96
- Tao goi thau: chi hoi Du an, Ten goi, Phuong an (PA1/2/3), Vi tri (Xa bien / Gan bien); khong can nap file khi tao - nap sau bang Nhap du lieu tu excel

## 0.9.97
- Doi "PA" thanh "Quy trinh": Du toan DNTC / NTP-BCHCT soat xet / NTP-BCHCT-P.KSCP (bo chu PA1 PA2 PA3); dau bang can lai cot cho vua

## 0.9.98
- Thanh ben: thay chu CCM-AI bang logo NEWTECONS (nen trang)

## 0.9.99
- Logo thanh ben thu gon (cao giam 1/2)

## 0.9.100
- Bang soat xet: moi phan mot mau (Quy trinh: Du toan, NTP/BCHCT, NTP/BCHCT/P.KSCP; Phan 1 do nhat, Phan 2 xanh duong nhat, Phan 3 cam nhat); duong ke doc ngan cach cac phan chay tu tieu de xuong het bang

## 0.9.101
- Giao dien v2 theo mau de xuat (chi doi kieu hien thi, khong doi noi dung / chuc nang): mau NEWTECONS, thanh ben gon, logo to day o, the trang bo goc, nut / bang / tab dong bo; them file frontend/src/ccm/theme-v2.css
