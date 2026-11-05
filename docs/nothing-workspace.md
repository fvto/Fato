# Giao diện Chingluh lấy cảm hứng từ Nothing

Triển khai ngày 26/10/2026 trên repository hiện có. Cập nhật trang chính `/` và trang báo cáo `/reports` với nền trắng ngà/đen, tiêu đề serif, wordmark dạng chấm, thanh điều hướng nổi và menu dạng panel. Đây là thiết kế cho công cụ nội bộ Chingluh, dựa trên ngôn ngữ thị giác đã khảo sát của Nothing; không phải bản sao toàn bộ website thương mại điện tử.

## Thành phần và chuyển động

- Thanh điều hướng và menu dùng chung trong `templates/_nothing_shell.html`. Dialog hỗ trợ Escape, giữ focus trong menu và trả focus về nút mở sau khi đóng.
- Menu mở khoảng 340 ms, đóng khoảng 160 ms; chuyển nội dung tool khoảng 240 ms; phản hồi hover khoảng 200–280 ms. Phần giới thiệu xuất hiện nhẹ khi tải trang.
- Tôn trọng `prefers-reduced-motion`: bỏ chuyển động trang trí và chuyển panel tức thời khi người dùng yêu cầu giảm chuyển động.
- Thẻ tool vẫn bấm được ở tiêu đề, mô tả, hình minh họa và vùng nền. Nút yêu thích có thao tác riêng. Enter/Space tiếp tục chuyển tool.
- Menu có đường dẫn trực tiếp đến từng tool qua `/reports?tool=...`. Mở trang/chuyển tool không tự chạy báo cáo.
- Trạng thái job và tiến độ vẫn do backend cung cấp; animation không tạo phần trăm giả hoặc đổi trạng thái xử lý.
- Giao diện chính giữ lựa chọn nhiều nhà máy, điều khiển SFC, lệnh cũ, các hộp thoại, nhật ký và đổi sáng/tối. Chế độ sáng là mặc định của thiết kế mới.
- Khi chưa có ảnh SFC, hiển thị hướng dẫn kết nối thay cho ảnh lỗi.

## File thay đổi

- `templates/index.html`: bố trí phần mở đầu/thanh nhà máy, nạp skin mới, menu dùng chung và ảnh trống. Toàn bộ mã JavaScript nghiệp vụ inline cũ và các ID điều khiển cũ được giữ nguyên.
- `templates/report_tools.html`: bố cục trang báo cáo mới; giữ các ID mà luồng xử lý đang sử dụng.
- `templates/_nothing_shell.html`: điều hướng và dialog dùng chung.
- `static/workspace.css`, `static/workspace.js`: token màu/font, shell và animation trình bày. Script này không gửi request hoặc đọc dữ liệu file.
- `static/report-tools.css`: giao diện thẻ tool, cấu hình, upload, tiến độ và export center.
- `static/dashboard-theme.css`: skin của các điều khiển Dashboard hiện có, bao gồm sáng/tối.
- `static/report_tools.js`: hình thẻ, animation chuyển tool, hỗ trợ URL chọn tool; giữ luồng nguồn tự động và tương thích API cũ/mới.
- `static/chingluh-wordmark.svg`, `static/report-object.svg`: vector gốc cho Chingluh; các hình nhỏ của từng tool được dựng bằng SVG trong giao diện.
- `static/vendor/`: Source Serif 4 hỗ trợ tiếng Việt và bộ Font Awesome 6.4.0 vốn được ứng dụng dùng, nay phục vụ cục bộ. License và nguồn lưu cùng tài nguyên.
- `tests/test_workspace_appearance.py`: kiểm tra shell, menu, navigation, mobile, giảm chuyển động, nhà máy, theme và hộp thoại.
- `tests/test_report_browser.py`: mở rộng kiểm tra chuyển tool khi bấm hình minh họa.
- `tests/test_report_engines.py`: chờ worker thực sự thoát trước khi dọn workspace tạm trên Windows; tránh lỗi test khi kết quả đã lưu nhưng tiến trình chưa kết thúc.

## Tương thích và dữ liệu

Giữ chế độ **Tự động: dùng dữ liệu có sẵn, tải thiếu từ SFC**, cả bốn nguồn dữ liệu, các chế độ báo cáo và liên kết tải kết quả. Bộ xử lý, bộ lọc acquisition, format Excel/PPTX và API thực thi không bị thay đổi trong lần đổi giao diện này.

Font, icon và ảnh trang trí được phục vụ từ máy chạy Dashboard. Không có request tải font/icon bên ngoài khi sử dụng giao diện mới. Các tài nguyên tham khảo công khai được tải vào máy lúc phát triển; dữ liệu công ty không được gửi đi.

## Kiểm chứng

- Bộ Python: 111 test được phát hiện; 107 test qua, 4 test trình duyệt được chạy riêng.
- Hai test `test_report_browser.py` qua: chuyển tool bằng chuột/bàn phím, API cũ/mới, bốn nguồn, preset, validate, run, cancel, cảnh báo và download.
- Hai test `test_workspace_appearance.py` qua: menu/Escape/focus, đường dẫn từng tool, desktop 1440 px, mobile 390 px, giảm chuyển động, chọn nhà máy, hộp thoại, sáng/tối và không gọi tài nguyên ngoài.
- Các test JavaScript hiện có, kiểm tra cú pháp JavaScript và lint Python phần test thay đổi đều qua.
- Ba engine thật vẫn xuất Excel/PowerPoint/ZIP trên dữ liệu giả lập trong workspace tạm. Không chạy tác vụ tải SFC thật trong kiểm thử giao diện.

```powershell
python -m unittest discover -s tests -p "test_*.py"
node --check static/workspace.js
node --check static/report_tools.js
node tests/test_sfc_dashboard.js
node tests/test_sfc_export_ui.js
$env:CHINGLUH_UI_TESTS='1'
python -m unittest discover -s tests -p "test_report_browser.py"
python -m unittest discover -s tests -p "test_workspace_appearance.py"
```

Không có bước build frontend. Server hiện tại đã phục vụ template và static assets mới: mở `/` hoặc `/reports`, nhấn **Ctrl+F5**. Các kiểm tra thao tác SFC trong trình duyệt sử dụng phản hồi giả lập; trạng thái đăng nhập SFC thực tế vẫn phụ thuộc phiên của người dùng.
