# CCM-AI 0.2 — AI cục bộ

Ứng dụng gọi Ollama tại `127.0.0.1:11434`, dùng Qwen3 4B chạy CPU. Không cần API key. Script khởi động đặt `OLLAMA_NO_CLOUD=1`; backend chặn tên mô hình cloud và mô hình có remote_host/remote_model. Không gọi tìm kiếm web. Tải runtime/mô hình cần Internet; xử lý hồ sơ bằng mô hình đã tải chạy trên máy.

## Cài và dùng

1. Chạy `Cài AI cục bộ.cmd` trong thư mục cài CCM-AI. Runtime CPU khoảng 36 MB, mô hình khoảng 2.5 GB. Script tải từ bản Ollama chính thức v0.35.0, kiểm tra SHA256 của byte range CPU và chữ ký Ollama Inc.; không thay driver GPU.
2. Chạy shortcut **Khởi động CCM-AI**. Trong Excel mở bảng điều khiển CCM-AI, mục **AI cục bộ** → **Cài đặt & kiểm tra AI**. Bấm **Thử AI bằng dữ liệu minh họa** rồi **Lưu cấu hình AI**.
3. Lấy hai vùng báo giá/giá soát xét, xác nhận ánh xạ và bấm **Bắt đầu soát xét**. Chọn một dòng cần phân tích.
4. Chọn thêm tối đa 3 tài liệu và 15 giá thư viện nếu cần. Bấm **Chuẩn bị nội dung cho AI** và đọc nội dung xem trước. Mặc định chỉ gồm dòng hiện tại, tối đa 8 ứng viên đối chiếu và quy tắc nghiệp vụ đang lưu. Tài liệu chỉ đưa tối đa 3 đoạn mỗi tài liệu, có tên/phiên bản/hash/dòng trích; đây là chọn đoạn theo từ khóa, chưa phải tìm kiếm toàn văn bằng embeddings.
5. Xác nhận nội dung rồi bấm **Phân tích bằng AI cục bộ**. Bấm **Dùng gợi ý trong bảng kiểm tra** hoặc **Từ chối và lưu phản hồi**. Gợi ý chưa ghi vào workbook.
6. Kiểm tra từng dòng, loại vật tư/nhân công, thông số, ĐVT, VAT, thời điểm và phạm vi. Xác nhận ghép dòng và cách tính thành tiền. Dùng **Xem trước các ô thay đổi** rồi **Áp dụng vào workbook** như quy trình cũ.

## Phạm vi AI

AI gợi ý ghép dòng, lựa chọn nguồn và giải thích thiếu/sai khác điều kiện. Model chỉ trả ID nguồn; đơn giá được backend lấy lại từ thư viện hoặc dòng đối chiếu đã cung cấp. Nguồn bịa/JSON không đúng schema bị chặn. Kết quả luôn cần kiểm tra, trạng thái vàng và chưa xác nhận. Người dùng quyết định lỗi thực tế; giá cao hơn không tự động là lỗi.

AI chưa tự phân rã giá gộp, đặt hệ số, tính tỷ lệ hoặc chọn cơ sở tính thành tiền. Hồ sơ PDF quét cần OCR chưa hỗ trợ. Ghép ngữ nghĩa có thể sai, nhất là hãng/model và đơn vị; tên nguồn đúng không chứng minh nội dung tương đương.

## Phản hồi và “huấn luyện”

Kết quả, ngữ cảnh đã xử lý và quyết định dùng/từ chối được lưu trong SQLite cục bộ, kinds `ai-result` và `ai-feedback`. Bản này **chưa fine-tune hoặc tự huấn luyện** từ các quyết định đó. Nạp thêm hồ sơ/thư viện giá bổ sung căn cứ cho lần phân tích; không thay trọng số mô hình. Có thể phát triển tuyển chọn ví dụ đã duyệt và huấn luyện riêng khi đã có dữ liệu chuẩn và bộ kiểm thử nghiệp vụ.

## Vận hành

Mô hình lưu tại `data/models`; log tại `data/ollama-error.log`. CPU cần thời gian, lần đầu có thể vài phút. Khi mất kết nối, thao tác ghép/soát xét thủ công vẫn dùng được. Không có retry suy luận tự động; bản xem trước hết hạn sau 10 phút hoặc sau một lần phân tích.

Nếu một Ollama khác đã chạy tại cổng 11434, script giữ nguyên dịch vụ đó; backend vẫn kiểm tra mô hình tải cục bộ. Cấu hình cloud-off chỉ áp dụng tiến trình do script CCM-AI khởi chạy. Bản ZIP không kèm mô hình lớn hoặc dữ liệu người dùng; máy mới chạy script cài AI để tải riêng.

Nguồn kỹ thuật: [Ollama Windows](https://docs.ollama.com/windows), [Local-only mode](https://docs.ollama.com/faq), [Chat API](https://docs.ollama.com/api/chat), [Qwen3 4B](https://ollama.com/library/qwen3:4b).

## Bản 0.3 — Ghi nhớ nghiệp vụ đã duyệt

Trong mục AI cục bộ, mở **Ghi nhớ nghiệp vụ đã duyệt**. Sau khi sửa và xác nhận dòng trong bảng kiểm tra, nhập bài học 10–1500 ký tự, đánh dấu đã duyệt và bấm **Lưu ví dụ chuẩn đã duyệt**. Ví dụ được lưu trong SQLite (`ai-memory`), còn sau khi đóng ứng dụng/khởi động lại dịch vụ.

Lần chuẩn bị AI tiếp theo, nếu bật **Dùng ví dụ nghiệp vụ đã duyệt phù hợp**, ứng dụng tìm tối đa 3 ví dụ cùng nghiệp vụ, cùng ĐVT và có mô tả tương tự bằng từ khóa. Các ví dụ xuất hiện ở `approvedExamples` trong bản xem trước, có mã, ngày duyệt và phiên bản quy tắc. AI có thể dùng bài học để diễn giải và dẫn mã `memory:...`. Không đảm bảo mọi trường hợp tương tự đều được tìm hoặc AI luôn dùng đúng bài học; kiểm tra bản xem trước và kết quả.

Đây là nhớ bằng tra cứu ví dụ, chưa huấn luyện lại trọng số mô hình. Quyết định dùng/từ chối gợi ý không tự biến thành ví dụ chuẩn. Giá số trong ví dụ cũ không đưa vào ngữ cảnh AI để dùng làm giá mới; giá đề xuất vẫn chỉ lấy từ các nguồn hiện tại đã chọn. Quy tắc hiện tại được áp dụng; phiên bản quy tắc cũ chỉ là lịch sử.

Có thể bấm **Thu hồi ví dụ này** để ngừng sử dụng; lịch sử vẫn giữ trên máy. Bản xem trước có ví dụ đã thu hồi bị chặn khi phân tích. Để sửa bài học, thu hồi ví dụ cũ rồi lưu lại từ dòng đã chỉnh. Ghi nhớ giúp giữ bài học, không tự đảm bảo kết luận đúng.