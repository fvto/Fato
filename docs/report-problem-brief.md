# Diễn giải cảnh báo báo cáo — 27/10/2026

Trong **Xem kết quả / Export center**, mỗi dòng kiểm tra có cảnh báo hiển thị nút nhỏ **Chi tiết** kèm biểu tượng cảnh báo. Nội dung mặc định ẩn. Bấm để mở bảng nổi cạnh nút; bấm lại, bấm ra ngoài, nút đóng hoặc Esc để đóng. Bảng tự giới hạn theo màn hình và có cuộn nội dung trên desktop/mobile. Dòng kiểm tra đạt không có nút này.

Nội dung gồm số bản ghi cần xem lại, từng nhóm vấn đề, lý do bị gắn cờ, số dòng theo nhà máy/loại nghiệp vụ, hướng kiểm tra, tên file/sheet và vị trí Excel. Các ngưỡng bất thường được diễn giải là **cần xác minh**, không khẳng định nhập sai. Không suy đoán nguyên nhân khi quy tắc chưa có diễn giải.

## Phạm vi

- COPQ: nhóm theo trường thiếu và tổ hợp quy tắc vi phạm. Một dòng có cả trường thiếu và lỗi liên kết vẫn chỉ được đếm một lần. Phân biệt Site với Factory; nêu số tiền vượt ngưỡng và giờ làm/số lượng ở các ví dụ phù hợp.
- HFPA: phân biệt lot chỉ có ở QAStation, chỉ có ở MES410 và số lỗi khác nhau giữa hai nguồn.
- Các kiểm tra khác: nêu kiểm tra chưa đạt và giá trị trước/sau khi có. Cảnh báo engine vẫn được đưa vào kết quả.
- Không đổi quy tắc nghiệp vụ, số liệu nguồn, workbook kết quả hay điều kiện xác nhận tải file.
- Tóm tắt được tạo trong lần đọc workbook để đối chiếu hiện có, không đọc lại workbook ở mỗi lần tải trang. Lưu ở `export_review.brief`; chi tiết từng quy tắc ở `checks[].issues`.
- Mỗi nhóm lưu số lượng chính xác, tối đa 12 vị trí dòng và 3 ví dụ số liệu. Giao diện ghi rõ khi chỉ hiển thị các vị trí đầu tiên; vẫn xem toàn bộ các dòng trong Excel. Hiển thị 5 nhóm đầu và cho mở rộng các nhóm còn lại.
- Lượt cũ chưa có phân nhóm được tóm tắt từ các kiểm tra đã lưu, có thông báo giới hạn, không tự quét lại toàn bộ file kết quả. Lượt COPQ tháng 04/2026 với 238 dòng cảnh báo đã được bổ sung phân nhóm 234 + 2 + 1 + 1 sau khi xác nhận checksum các file.

## Xác minh

- Bộ Python: 124 test được phát hiện, 118 qua và 6 test trình duyệt tùy chọn bỏ qua.
- Chạy riêng `tests/test_report_browser.py` với `CHINGLUH_UI_TESTS=1`: cả 3 bài qua, gồm thao tác nút, đóng bằng click/Esc, nội dung, xác nhận tải, chống chèn HTML và màn hình rộng 1280/390 px.
- Đã xem ảnh chụp desktop/mobile, kiểm tra cú pháp JavaScript và Python.
