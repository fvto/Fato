# Tối ưu kiểm tra đầu vào — 27/10/2026

Áp dụng cho nút **Kiểm tra đầu vào** tại `/reports`, dùng chung cho COPQ, HFPA và Re-Inspection.

- Bước kiểm tra chỉ tạo bản sao đầu vào; không sao chép mã nguồn/template của tool và không ghi lại workbook đã chuẩn hóa. Chuẩn hóa vẫn thực hiện khi chạy báo cáo.
- Nếu chỉ tải file từ máy, không quét danh sách file local/SFC. Nguồn tự động dùng lại danh sách mà bước lập kế hoạch vừa chọn, tránh quét lần thứ hai.
- Lưu tối đa 64 kết quả kiểm tra trong bộ nhớ tiến trình server. Mỗi lần kiểm tra vẫn đọc SHA-256 toàn bộ bản sao đầu vào. Chỉ dùng lại kết quả khi nội dung, loại file, tool, nhóm, tháng, MIME và hợp đồng dữ liệu đều khớp. Khởi động lại server sẽ xóa bộ nhớ này.
- Không lưu các dòng dữ liệu trong bộ nhớ đệm. Mỗi phản hồi nhận bản sao riêng của kết quả; đổi tên file sẽ hiển thị tên mới. File lỗi cũng có thể dùng lại kết quả nếu nội dung/quy tắc không đổi.
- Luồng chạy vẫn kiểm tra toàn bộ dữ liệu độc lập, chuẩn hóa bản sao và kiểm tra lại trong worker. Các quy tắc báo thiếu, sai định dạng, trùng tên file/cột và giới hạn hiện có được giữ nguyên.

API `/validate` trả thêm `checked_files`, `cached_files` và `validation_time_ms`; client hiện có vẫn tương thích.

## Đo trên dữ liệu giả lập

Một workbook COPQ gồm 10.000 dòng, 6 cột, tải qua `FileStorage`; không kết nối SFC. Cùng tiến trình và cùng nội dung:

| Phép đo | Thời gian |
| --- | ---: |
| Kiểm tra có ghi workbook chuẩn hóa như trước | 1,1086 giây |
| Kiểm tra chỉ đọc lần đầu | 0,7570 giây |
| Kiểm tra lại nội dung không đổi, lần 1 | 0,0311 giây |
| Kiểm tra lại nội dung không đổi, lần 2 | 0,0111 giây |

Phép so sánh chuẩn hóa dùng cùng hàm gọi mới, thay bộ kiểm tra bằng `validate_file(normalize=True)`; không tính chi phí sao chép bộ tool cũ. Đây là phép đo cục bộ trên file giả lập; thời gian thực tế phụ thuộc số file, định dạng, dung lượng, ổ đĩa và bước tìm nguồn tự động.

## Xác minh

`python -m unittest discover -s tests -p 'test_*.py'`: 117 test, 112 qua và 5 test trình duyệt tùy chọn bỏ qua.

Năm test mới tại `tests/test_report_validation_speed.py` xác minh dùng lại nội dung, không chia sẻ dữ liệu kết quả có thể sửa, phát hiện thay đổi dù cùng dung lượng/thời gian sửa, đổi tháng/MIME/schema/hợp đồng, giới hạn bộ nhớ đệm và API upload không sao chép/ghi workbook/quét danh sách không cần thiết. Các test độ tin cậy hiện có tiếp tục xác minh lỗi ở dòng cuối và API chạy không bỏ qua kiểm tra đầu vào.

Đã nạp lại server tại cổng 5050 sau khi xác nhận không có báo cáo, lệnh SFC hoặc MQAA đang chạy. Gọi API HTTP thật hai lần bằng CSV giả lập: đều hợp lệ, lần đầu `cached_files=0`, lần sau `cached_files=1`. Không chạy báo cáo hoặc tải dữ liệu SFC trong phép kiểm tra này.
