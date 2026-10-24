# Chingluh Dashboard

Mở `MO_BANG_DIEU_KHIEN.bat`:

1. **Mở Dashboard**: chạy server và mở giao diện.
2. **Mở dashboard và DOM**: chạy server, mở kết nối SFC và giao diện. Đăng nhập SFC nếu được yêu cầu.

`MO_SERVER.bat` chạy riêng server tại http://localhost:5050/.

- Code Dashboard nằm trong thư mục này; file khởi động Python là `dashboard_launcher.py`.
- Kết quả HFPA và các phần tải để tiếp tục khi lỗi nằm trong `Report/HFPAExports`.
- Các bài kiểm tra nằm trong `tests`; nhật ký và công cụ chẩn đoán nằm trong `scratch`.
- Python, Node.js và Chrome/Edge sử dụng môi trường đã có trên máy.

## Kết nối MQAA cũ

Các chương trình quét, dữ liệu đối chiếu và file Master MQAA vẫn nằm trong thư mục `Documents/MQAA-Record`.
Dashboard gọi chương trình MQAA bằng tiến trình riêng và đọc file Master qua đường dẫn trong `dashboard_config.json`.
Nếu thư mục MQAA đổi vị trí, cập nhật `legacy_mqaa_root` trong file cấu hình.

`migration_manifest.json` ghi lại các file đã chuyển và mã kiểm tra nội dung Excel khi tách thư mục.
