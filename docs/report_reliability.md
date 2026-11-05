# Audit và cải thiện tool báo cáo — 26/10/2026

## Hiện trạng đã xác minh

- Repository hiện có: `C:\Users\User\Documents\Chingluh Dashboard`. Python/Flask, giao diện HTML/JavaScript thuần; không có React/Vite/npm build.
- `dashboard_launcher.py`, hai BAT, `dashboard_paths.py`: khởi động và phân giải đường dẫn. MQAA dùng workspace cũ cấu hình riêng; không di chuyển/chỉnh file Master ngoài repo.
- `templates/index.html`, `sfc_controller.py`, `sfc_http.py`, `sfc_downloads.py`, `sfc_excel.py`: giao diện SFC, trình duyệt/HTTP, tải và ghép file. Luồng lệnh đã có SSE và lưu kết quả.
- `sfc_report_tools.py`, `report_tool_worker.py`, `sfc_report_acquisition.py` và ba profile acquisition: chọn nguồn, thu thập đúng bộ lọc, workspace riêng và worker Python tách biệt.
- `report_tools/reinspection`, `copq`, `hfpa`: bộ xử lý và template Excel/PowerPoint thực tế. Các module mang tên `agents` là mã xử lý cục bộ được repo sử dụng.
- `Report/HFPAExports` lưu file SFC; `Report/ToolRuns/<id>` lưu bản sao và kết quả. Đã có lịch sử chạy, chọn nhiều file theo nhóm/nhà máy, tải từng file/ZIP, dùng lại dữ liệu và khôi phục bộ lọc.
- Trước thay đổi: 86 test Python qua. Có nhiều thay đổi/untracked files từ trước; không reset, checkout hoặc ghi đè các thay đổi đó.

## Gap P0 và cách xử lý

### Đầu vào

Trước đây đường thủ công chỉ kiểm tra đuôi/tên/tháng, trong khi acquisition có kiểm tra sâu hơn nhưng không đồng nhất. Các engine có nhánh coercion/skip nên lỗi có thể trở thành 0 hoặc bị bỏ khỏi tổng hợp.

`report_validation.py` dùng `report_contracts.json`: nhận diện chữ ký OOXML/OLE/CSV, giới hạn mặc định 50 MB/file, 256 MB khi giải nén, 500.000 dòng và 2.048 cột/sheet. Có thể đặt `limits` riêng trong schema. Tổng multipart giới hạn 640 MB. MIME là tín hiệu phụ; nội dung thật luôn được đọc. OOXML mang đuôi `.xls` của SFC tiếp tục được hỗ trợ.

Kiểm tra tất cả bản ghi, lỗi có vị trí và hướng sửa, không in giá trị ô. Kiểm tra các cột bắt buộc, trùng tiêu đề, dữ liệu ngoài tiêu đề, ngày/kỳ, số không hợp lệ/âm, công thức đầu vào và HFPA Station. Các số đếm dùng kiểu nguyên khi engine yêu cầu; COPQ vẫn nhận số đôi lẻ, chỉ số sản lượng và nhân lực vẫn nhận số lẻ hợp lệ. Chuỗi số có phân nhóm nghìn chuẩn được đọc thành số.

Những tương thích quan trọng: hai cột `Defective Rate` trong Re-Inspection được giữ; Audit Date gốc Re-Inspection có thể ngoài kỳ hoặc trống; các alias cũ đã biết được đổi thành tên chuẩn trong bản sao. Không gộp lẫn sheet khi engine chỉ đọc sheet đầu. COPQ/FTT theo cách đọc nhiều sheet/đầu vào của từng engine. Chỉ bản sao được chuẩn hóa thành `.xlsx`, bỏ macro/liên kết ngoài; SHA-256 đầu vào ghi nhận file nguồn.

Dòng tổng DevExpress chỉ được loại nếu là dòng cuối có dữ liệu, chỉ chứa trường đếm/tổng phù hợp và số đếm bằng số bản ghi đã đọc. Vị trí loại được ghi rõ; dòng không khớp vẫn được kiểm tra và chặn khi lỗi. Dòng trắng có số lượng riêng.

### Thực thi

`report_runtime.py`: danh sách bước pending/running/done/error/cancelled, thời điểm bắt đầu/kết thúc, dữ liệu lưu nguyên tử; thử lại ngắn khi Windows tạm khóa file trạng thái. Poll đọc lại trạng thái cuối trước khi kết luận worker chết để tránh ghi đè kết quả vừa hoàn tất.

Cancellation dùng `cancel.request`, kiểm tra ở ranh giới bước/file. Đang tải hoặc ghi thì chờ kết thúc, không kill tiến trình. Thu thập SFC giữ file tải xong và chạy phần khôi phục bộ lọc; timeout/mất kết nối vẫn giữ quy tắc dừng cũ. Không thêm tải lại tự động. Số file là tiến độ đo được; không hiển thị phần trăm giả.

Output thô của engine bị chặn khỏi log kỹ thuật. Lỗi UI có hướng xử lý; diagnostics chỉ có loại lỗi, tên file mã nguồn, hàm và dòng. Legacy logs không được đưa thẳng ra API. Báo cáo nghiệp vụ tạo ra vẫn là dữ liệu nội bộ, cần quản lý quyền truy cập ổ đĩa như trước.

### Xuất file

`report_exports.py` đối chiếu:

- COPQ: số dòng nguồn so với master sau khi loại Rework theo hành vi cũ, tổng chi phí với dung sai 0,01, cột trạng thái chất lượng của engine.
- FTT B/C: số dòng và số lượng lỗi sau bộ lọc B/C của engine.
- HFPA: số dòng QAStation/MES410 trong master và trạng thái lot trong `Lot_Defect_Variance`.
- Re-Inspection: số bản ghi/QC/MA mà engine đã đọc theo nhà máy, đủ sheet cần có, tổng QC trong database theo mẫu hiện hành.
- Mỗi mode có định dạng bắt buộc khai báo trong registry; thiếu bảng đối chiếu hoặc file chưa hoàn chỉnh luôn hiện cảnh báo/lỗi, không coi là qua.

File có metadata dung lượng, thời điểm sửa, SHA-256, sheet và số dòng vật lý của sheet (bao gồm tiêu đề). ZIP giữ cấu trúc liên kết cũ. File Excel/PowerPoint không đọc được không được công bố. Kết quả dở dang/có sai lệch yêu cầu xác nhận; kiểm tra lại checksum tại download để phát hiện thay đổi sau đối chiếu. Không tự sửa số liệu để ép khớp.

## API và tương thích

- API chạy cũ `POST /api/report-tools/<tool>/runs` tiếp tục hoạt động và luôn preflight phía server. Lỗi schema: HTTP 422 kèm `validation`; lỗi cấu hình: 400; bận: 409.
- `POST /api/report-tools/<tool>/validate` kiểm tra trên workspace tạm, không spawn worker hoặc mở SFC. Nguồn auto chỉ lập kế hoạch; danh sách còn thiếu được trả về, sau đó kiểm tra lại toàn bộ file khi thu thập xong.
- `POST /api/report-tools/runs/<id>/cancel` yêu cầu hủy an toàn. Không nhận dạng PID cũ để cưỡng bức kill.
- API danh sách cũ giữ nguyên; UI dùng `GET /api/report-tools?catalog=0` để tải metadata nhanh và `/<tool>/inputs?source=local|sfc` khi cần danh sách file.
- Download cũ giữ đường dẫn; lượt mới cần `?ack=1` khi `requires_ack`. Lượt cũ chưa có checksum giữ hành vi tải trước đây.
- UI dùng registry cho thẻ, mode, nhóm, định dạng; preset và favorites chỉ lưu cài đặt trong localStorage. Thêm tool bằng metadata, schema và adapter `module:function` cục bộ; nghiệp vụ mới vẫn phải có kiểm tra đối chiếu và test phù hợp.

## P1/P2 đã làm và chưa ưu tiên

Đã thêm search, recent/favorites, drop zone, preset theo tool, luồng bốn bước, bảng trước/sau theo các phép đối chiếu. Tái sử dụng lịch sử chạy hiện có. Metadata/config tách khỏi UI; COPQ có callback bước mới nhưng tham số tùy chọn nên CLI cũ vẫn gọi được.

Chưa thêm smart detection theo cấu trúc trên toàn thư mục, hàng đợi chạy nhiều lượt, audit log người dùng nhiều tài khoản, so sánh hai phiên bất kỳ hoặc workflow chaining. Xử lý nhiều file theo nhà máy và thu thập theo lô vốn đã có, không xây lại thành hệ thống mới.

## Bảo mật và retention

Mọi xử lý mới nằm cục bộ. Không thêm upload cloud/API AI. Trang report không tải tài nguyên bên ngoài. Mặc định server bind loopback; có biến môi trường `CHINGLUH_BIND_HOST` để giữ khả năng triển khai LAN có chủ đích. POST report từ Origin khác bị từ chối; response report có no-store/nosniff. Tên file tải lên/tải xuống được làm sạch, kiểm tra đường dẫn/symlink output.

`report_retention.py --days 30` chỉ liệt kê lượt mới đã kết thúc và hết hạn. `--apply` xóa workspace và kết quả các lượt đó sau khi kiểm tra lại phạm vi đường dẫn/trạng thái. Không có tác vụ xóa tự động; không đụng lượt legacy, đang chạy, dữ liệu Desktop hoặc cache SFC.

## Cách chạy và kiểm chứng

Khởi động Dashboard bằng BAT hiện có, mở `http://localhost:5050/reports`. Nếu server cũ còn chạy, cần khởi động lại tiến trình server để nạp API mới; tải lại trang chỉ nạp giao diện.

```powershell
python -m unittest discover -s tests -p "test_*.py"
node tests/test_sfc_dashboard.js
node tests/test_sfc_export_ui.js
node --check static/report_tools.js
python -m compileall -q report_validation.py report_runtime.py report_exports.py report_retention.py sfc_report_tools.py report_tool_worker.py sfc_report_acquisition.py sfc_dashboard_server.py
$env:CHINGLUH_UI_TESTS='1'
python -m unittest discover -s tests -p "test_report_browser.py"
```

Lint dùng Ruff 0.14.2 cho các module/test mới và phần report đã sửa; kiểm tra F-rules trên server/adapter cũ. Không chạy format toàn bộ server để tránh trộn thay đổi không liên quan. Bài JS cũ được cập nhật fixture/expectation cho route `/api/action` mà form hiện tại đã dùng trước thay đổi; không đổi hành vi form gốc.

Các test bao gồm subprocess thật của ba engine tạo Excel/PPTX/ZIP, hủy trước worker và sau download, khôi phục SFC bằng controller giả, lỗi Windows khi ghi trạng thái, poll trùng thời điểm worker hoàn tất, chặn tải khi chưa xác nhận hoặc checksum thay đổi, retention, dữ liệu lỗi ở cuối file và kiểm tra trình duyệt desktop/mobile.

Chưa thực hiện tải mới từ phiên SFC đăng nhập thật, thử mất mạng/Excel khóa file bằng ứng dụng thật, kiểm tra trực quan từng slide/bảng và kiểm thử mọi mẫu lịch sử/tháng hoàn toàn rỗng hoặc dữ liệu sát giới hạn dung lượng. Mẫu SFC có sẵn được đọc cục bộ để đối chiếu schema; không chỉnh file gốc. Đối chiếu tự động không thay thế phê duyệt ý nghĩa nghiệp vụ của báo cáo.

## File thay đổi trong lần triển khai này

- Backend hiện có: `sfc_report_tools.py`, `report_tool_worker.py`, `sfc_report_acquisition.py`, `sfc_dashboard_server.py`.
- Adapter engine: `report_tools/copq/agents/pipeline_orchestrator.py` (callback bước tùy chọn; CLI cũ giữ nguyên).
- Module/config mới: `report_validation.py`, `report_runtime.py`, `report_exports.py`, `report_retention.py`, `report_contracts.json`, `report_registry.json`.
- Giao diện: `templates/report_tools.html`, `static/report_tools.js`.
- Test mới: `tests/test_report_reliability.py`, `tests/test_report_engines.py`, `tests/test_report_browser.py`.
- Test fixture cập nhật: `tests/test_report_tools.py`, `tests/test_sfc_dashboard.js`.
- Tài liệu: `README.md`, `docs/report_reliability.md`.

Không sửa `templates/index.html`, engine HFPA/Re-Inspection, template báo cáo, bộ lọc SFC hoặc config nguồn Desktop trong lần triển khai này. Những file đã có thay đổi từ trước vẫn được giữ nguyên.

## Trạng thái đưa vào sử dụng

Kết quả kiểm thử cuối ngày 26/10/2026:

- `unittest discover`: 108 test được phát hiện, 107 qua, 1 test Chrome được tách riêng; tổng thời gian 22,797 giây.
- Test Chrome chạy riêng: qua, gồm upload/validate/run/cancel/review/download, preset/favorites/search và viewport 390 px. Không có lỗi JavaScript trên trang.
- Hai bộ test JavaScript của trang chính: qua. Kiểm tra cú pháp JavaScript và biên dịch cú pháp các module Python: qua.
- Ruff 0.14.2: qua cho phần report/test đã triển khai; kiểm tra F-rules server và adapter cũ: qua.
- Cả ba engine thật xuất Excel, PowerPoint và ZIP trên dữ liệu giả lập; Quality Sentinel COPQ chạy trong workspace tạm: qua.
- Kiểm tra schema toàn bộ dòng trên các mẫu cục bộ đã đọc; mẫu COPQ legacy 1.143 bản ghi và 1 dòng tổng `Count=...` được xác nhận hợp lệ. Không lưu giá trị ô vào báo cáo kiểm thử này.

Code đã lưu trong repo. Server cũ tại cổng 5050 chưa được nạp lại: công cụ xét duyệt tự động từ chối lệnh dừng/khởi động lại với thông báo `blocked by policy`. Chưa có tiến trình nào bị dừng bởi lệnh đó.

Sau khi không còn tác vụ đang chạy, đóng tiến trình server Dashboard cũ bằng cách quản trị thường dùng trên máy, rồi mở `MO_SERVER.bat` hoặc `MO_BANG_DIEU_KHIEN.bat` và tải lại `/reports`. Chỉ tải lại trang hoặc bấm BAT khi server cũ vẫn chạy sẽ chưa nạp API Python mới. Các API validation/monitor mới cần server mới.

### Khôi phục nguồn tự động theo yêu cầu

Chế độ mặc định tiếp tục là **Tự động: dùng dữ liệu có sẵn, tải thiếu từ SFC**. UI nhận diện được API cũ vốn chưa trả về các trường `automatic`, `contract_version`, `required_groups`, để lựa chọn tự động không bị ẩn nhầm. Với server cũ, nguồn thư mục/SFC dùng danh sách file đã có trong API danh sách cũ và nút Chạy sử dụng luồng chạy/kiểm tra acquisition cũ; không gọi API validation mới chưa tồn tại. Khi server đã nạp bản mới, bước validation và monitor mới hoạt động bình thường.

Khôi phục thẻ mô tả nguồn tự động/bộ lọc theo từng tool, giữ đủ bốn lựa chọn nguồn. File đã có và hợp lệ đúng kỳ được dùng lại, chỉ tải các nhà máy/nhóm còn thiếu từ SFC. Kiểm thử Chrome bao gồm cả metadata mới và metadata cũ, danh sách local/SFC, đủ bốn nguồn và payload `acquire_missing=true`; không tải file công ty thật trong kiểm thử. Phần sửa giao diện này chỉ cần tải lại trang bằng Ctrl+F5.
