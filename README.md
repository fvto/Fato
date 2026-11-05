# Chingluh Dashboard

Giao diện hiện tại dùng phong cách tối giản lấy cảm hứng từ Nothing trên cả trang chính và `/reports`: nền trắng ngà, tiêu đề serif, wordmark dạng chấm, menu nổi và chuyển động có hỗ trợ giảm motion. Nhấn **Ctrl+F5** để tải giao diện mới. Font/icon được phục vụ cục bộ. Xem `docs/nothing-workspace.md` cho danh sách thay đổi và kiểm thử.

Mở `MO_BANG_DIEU_KHIEN.bat`:

1. **Mở Dashboard**: chạy server và mở giao diện.
2. **Mở dashboard và DOM**: chạy server, mở kết nối SFC và giao diện. Đăng nhập SFC nếu được yêu cầu.

`MO_SERVER.bat` chạy riêng server tại http://localhost:5050/.

**Combine Excel:** mở `http://localhost:5050/combine` hoặc bấm **Combine Excel** trong Dashboard. Dùng chung cho bảng bán hàng, tồn kho, sản xuất, chất lượng và các bảng Excel có một dòng tiêu đề. Chọn/kéo thả `.xls`, `.xlsx`, `.xlsm`, kiểm tra trang tính, dòng tiêu đề được gợi ý và dòng cuối rồi **Xem trước → Gộp & tạo file Excel → Tải file**. Có thể giữ mọi cột (ô thiếu để trống) hoặc chỉ gộp khi cùng bộ cột, kể cả khác thứ tự. Tên cột được chuẩn hóa Unicode, khoảng trắng và hoa thường; không tự ghép cột khác tên/đơn vị. Chỉ bỏ cột trống nếu dữ liệu bên dưới cũng trống. Sheet `Sources` ghi phạm vi, số dòng và cột thiếu của từng bảng. Mẫu xuất đặc thù được nhận diện theo cấu trúc trong `spreadsheet_combine_profiles.py`; dòng tổng chỉ bỏ sau đối chiếu, có thể chọn giữ. Với mẫu chưa nhận diện, dùng dòng cuối để bỏ phần tổng/ghi chú. Gộp lại bản đã xuất vẫn giữ cột nguồn cũ, thêm cột truy vết mới với tên riêng. Công thức dùng kết quả đã lưu; không chạy công thức hoặc macro. Giới hạn: **50 file/lượt**, 40 MB/file, 500.000 ô/lượt. Dữ liệu nằm trong `Report/CombinedExcel`. Bộ xuất dùng runtime Artifact Tool của Codex; cấu hình đường dẫn qua `CHINGLUH_ARTIFACT_NODE` và `CHINGLUH_ARTIFACT_MODULES`.

**Spreadsheet Chat:** mở `http://localhost:5050/compare`, tải file rồi gõ yêu cầu bằng tiếng Việt hoặc tiếng Anh. Excel MCP Server và Codex CLI tự đọc cấu trúc file, kiểm tra phép tính hoặc đối chiếu dữ liệu và trả bằng chứng theo ô; có thể hỏi tiếp trong cùng cuộc chat. Cài runtime bằng `python scripts/setup_excel_ai.py` và đăng nhập Codex CLI. Nội dung liên quan được gửi đến AI qua tài khoản Codex đang đăng nhập. Chế độ chọn quy tắc thủ công vẫn có trong Advanced. Xem `docs/spreadsheet-compare.md`.

- Code Dashboard nằm trong thư mục này; file khởi động Python là `dashboard_launcher.py`.
- Kết quả HFPA và các phần tải để tiếp tục khi lỗi nằm trong `Report/HFPAExports`.
- Các bài kiểm tra nằm trong `tests`; nhật ký và công cụ chẩn đoán nằm trong `scratch`.
- Python, Node.js và Chrome/Edge sử dụng môi trường đã có trên máy.

Tên file xuất được suy ra từ phần tên chung và kỳ của các nguồn đã chọn, không gắn riêng với một loại báo cáo. Nhận ngày `YYYY-MM-DD`, `YYYYMMDD`, tháng `YYYY-MM` và dấu phân cách `_`/`.`. Ví dụ: `Sales_2026-09_2026-10_Combined.xlsx`, `Stock_2026-10-02_2026-10-08_Combined.xlsx`, `Payroll_2files_Combined.xlsx`. Khi kỳ không rõ hoặc không cùng mức ngày/tháng, dùng tên chung/trang tính và số file, không tự thêm ngày. Có thể sửa **Tên file xuất**; tên lưu và tải xuống thống nhất, luôn có đuôi `.xlsx`. File gộp cũ vẫn tải được.

## Kết nối MQAA cũ

Các chương trình quét, dữ liệu đối chiếu và file Master MQAA vẫn nằm trong thư mục `Documents/MQAA-Record`.
Dashboard gọi chương trình MQAA bằng tiến trình riêng và đọc file Master qua đường dẫn trong `dashboard_config.json`.
Nếu thư mục MQAA đổi vị trí, cập nhật `legacy_mqaa_root` trong file cấu hình.

`migration_manifest.json` ghi lại các file đã chuyển và mã kiểm tra nội dung Excel khi tách thư mục.

## MQAA: tự tải theo ngày khi SFC báo SQL timeout

Khi xuất Excel trên `QA/MQAAReport.aspx` gặp **Execution Timeout Expired**, Dashboard chờ phản hồi lỗi kết thúc rồi tự tải lần lượt từng ngày trong khoảng đã chọn, giữ nguyên nhà máy và bộ lọc. Mẫu **All report** được ghép thành một bảng Excel, kèm sheet `NguonTheoNgay` để đối chiếu file và số dòng từng ngày. Các mẫu phân tích tổng hợp được giữ riêng từng sheet ngày trong cùng một workbook.

Chỉ công bố file gộp khi đủ tất cả ngày, đúng cấu trúc cột và đúng nhà máy/ngày dữ liệu. Nếu một ngày tiếp tục lỗi, giữ các file ngày đã tải và báo rõ ngày lỗi; không tạo file gộp thiếu ngày. Lỗi quyền truy cập và timeout phía trình duyệt khi chưa có phản hồi máy chủ không kích hoạt tải lại. File nguồn và manifest nằm trong `Report/HFPAExports/`.

## Ba tool báo cáo

Trong Dashboard, bấm **Mở công cụ báo cáo**, hoặc mở `http://localhost:5050/reports`.

**HFPA & FTT theo quý** xuất cả Excel và PowerPoint trong một lượt, theo hai mẫu gốc `Quarterly Report HFPA-FTT.xlsx` và `Partner QBR Format Q2 FY27-CLG.pptx` trong Downloads. Nếu di chuyển mẫu, đặt đường dẫn trong `quarterly_templates.excel` và `quarterly_templates.powerpoint` của `report_tools_config.json`. Chọn quý đã có cột tương ứng trong mẫu Excel.

- **Tự động từ SFC**. Chọn quý và ngày chốt, bấm **Kiểm tra đầu vào** rồi **Chạy tool**. Tool lấy MES410 nội bộ và FTT Internal cho VH, VH2, JV, JV2. Báo cáo ghi ngày chốt và so sánh với phần tương ứng của quý trước. Nếu SFC yêu cầu đăng nhập, dùng Chrome kết nối với Dashboard.
- **Chọn file từ máy**. Sửa và lưu workbook theo mẫu, gồm số liệu và bảng Top model / Top lỗi của quý đã chọn. Tải file lên rồi bấm **Kiểm tra đầu vào → Chạy tool** để xuất cả hai bản cập nhật.

Với các công cụ báo cáo tháng:

1. Chọn **Re-Inspection**, **combine_COPQ** hoặc **HFPA**.
2. Chọn tháng, chế độ báo cáo và nguồn đầu vào: thư mục tool hiện có, các file đã tải từ SFC, hoặc tải Excel/CSV từ máy.
3. Mỗi nhóm dữ liệu chọn một file cho từng nhà máy trong tháng. Re-Inspection/HFPA cần VH, VH2, JV, JV2; COPQ hỗ trợ thêm VH3, VH4, JVB/JV3.
4. Bấm **Chạy báo cáo**. Xem tiến độ, nhật ký và tải từng kết quả hoặc gói ZIP giữ cấu trúc thư mục của các file liên kết.

Mã nguồn và mẫu đã tích hợp nằm trong `report_tools/`; `provenance.json` ghi nguồn và SHA-256. `report_tools_config.json` chỉ cấu hình nơi đọc dữ liệu cũ trên Desktop; có thể dùng đầu vào SFC hoặc tải file lên nếu các thư mục cũ đã di chuyển.

Mỗi lượt có bản sao mã nguồn, mẫu và đầu vào tại `Report/ToolRuns/<id>/workspace/`. Các bộ xử lý chỉ cập nhật bản sao này. File gốc trên Desktop và mẫu trong `report_tools/` được giữ nguyên.

Tác vụ báo cáo chạy bằng tiến trình Python riêng, tiếp tục khi đóng/tải lại trang. Mở lại trang chỉ xem trạng thái, không tự chạy lại. Mỗi thời điểm chạy một lượt báo cáo; các lượt bị lỗi vẫn giữ nhật ký và những file đã hoàn tất. Thư viện nằm trong `requirements.txt`.

### COPQ: chọn tháng và chạy hoàn toàn trên web

Chọn **combine_COPQ → Tự động: dùng dữ liệu có sẵn, tải thiếu từ SFC**, chọn tháng rồi bấm **Chạy báo cáo**. Không cần chạy BAT của tool hoặc chép file vào Input.

- COPQ lấy từ `QA/QAReportCOPQ.aspx`: xóa Category và Type; lấy toàn bộ tháng cho VH, VH2, VH3, VH4, JV, JV2, JVB/JV3.
- FTT lấy từ `QA/QA.aspx`: mở Select Station, bỏ chọn toàn bộ trạm và bấm **Internal**; chỉ lấy VH, VH2, JV, JV2, cùng kỳ tháng COPQ.
- File đúng mẫu, đủ kỳ và đúng nhà máy được dùng lại, kể cả file hợp lệ không có dòng dữ liệu. File thiếu mới được tải. Chế độ chỉ gộp COPQ không tải FTT; chế độ FTT B/C chỉ tải bốn file FTT.

### HFPA: hai nguồn tự động cùng tháng

Chọn **HFPA → Tự động: dùng dữ liệu có sẵn, tải thiếu từ SFC**, chọn tháng và bấm **Chạy báo cáo**. Cả hai nguồn đều lấy VH, VH2, JV, JV2:

- QAStation/FTT: `QA/QA.aspx` → chỉ chọn trạm **HFPA** → bấm **OK**.
- HFPA/MES410: `QA/QCHFPA.aspx?tree_id=Q504` → **tắt NIKE**, xóa lọc chuyền, tìm kiếm và xuất Excel nội bộ.
- Dùng lại file hợp lệ đúng tháng; kiểm tra toàn bộ dòng QAStation chỉ thuộc trạm HFPA, MES410 có cột nội bộ và Audit Date đúng kỳ. File MES410 gốc có tên theo thời điểm tải được nhận diện theo ngày kiểm tra bên trong.
- Chỉ tải các file còn thiếu, không tự gửi lại lệnh bị timeout. Không chạy báo cáo khi còn thiếu nhà máy; tải lại trang chỉ xem tiến độ/kết quả. Không cần BAT hay chép file vào Input.
- Nếu chưa có Chrome kết nối, hệ thống mở phiên SFC; người dùng đăng nhập khi được yêu cầu. Trong lúc thu thập, các thao tác SFC khác được chặn để giữ đúng bộ lọc. Sau khi đủ đầu vào, hệ thống trả trình duyệt về trạng thái trước đó và chạy báo cáo.
- Mất kết nối hoặc tải quá thời gian thì dừng, giữ các file đã hoàn tất và báo lỗi; không tự tải lặp lại. Tải lại trang chỉ xem tiến độ.

### Re-Inspection: ba nguồn tự động cùng tháng

Chọn **Re-Inspection → Tự động: dùng dữ liệu có sẵn, tải thiếu từ SFC**, chọn tháng rồi bấm **Chạy báo cáo**. Cả ba nguồn đều lấy VH, VH2, JV, JV2:

- Re-Inspection: `QA/QCRECHK.aspx` → để trống **Line** và **Type**, chọn đủ kỳ tháng, tìm kiếm rồi xuất Excel. Kiểm tra lại bộ lọc sau phản hồi SFC trước khi xuất.
- FTT: `QA/QA.aspx` → **không chọn trạm** → **Internal**, giống nguồn FTT của COPQ.
- HFPA: `QA/QCHFPA.aspx?tree_id=Q504` → **tắt NIKE**, xuất nội bộ; có thể dùng lại file hợp lệ đã tải cho tool HFPA.
- Cần đủ 12 file, mỗi nguồn một file cho từng nhà máy; dùng lại dữ liệu đúng mẫu và kỳ xuất, chỉ tải phần thiếu. `Audit Date` của Re-Inspection là ngày kiểm tra ban đầu và có thể nằm trước kỳ xuất; giữ nguyên các dòng SFC trả về. Dòng tổng cuối bảng không được nhầm với ngày kiểm tra.
- Database, Recycle Report và PowerPoint được tạo trong bản sao riêng. Tải kết quả ngay trên web; tải lại trang không khởi động lại tác vụ và không cần chạy BAT hay chép Input.

## Kiểm tra, theo dõi và xuất kết quả an toàn

Tại `/reports`, dùng luồng **Chọn dữ liệu → Kiểm tra đầu vào → Chạy tool → Xem và tải kết quả**. Bước kiểm tra chỉ đọc bản sao cục bộ, không khởi động SFC hay chạy báo cáo. API chạy cũ vẫn kiểm tra lại đầu vào, nên không thể bỏ qua kiểm tra bằng cách gọi trực tiếp.

- Kiểm tra chữ ký Excel/CSV, MIME tải lên, giới hạn dung lượng, sheet, cột, toàn bộ dòng, ngày và số liệu theo từng nguồn. Lỗi có sheet/dòng/cột và cách sửa; hiển thị tối đa 100 vị trí đầu tiên cùng tổng số lỗi. Không thay giá trị sai bằng 0. Dòng tổng DevExpress chỉ tách khỏi bản sao sau khi xác nhận cấu trúc và số bản ghi, có thông báo rõ.
- Theo dõi từng bước; số file đã tải/đã kiểm tra là số thực tế, không quy đổi thành phần trăm ước lượng. **Hủy tại điểm an toàn** chờ lời gọi tải/ghi hiện tại kết thúc; giữ file hoàn tất và khôi phục bộ lọc SFC khi kết nối còn hoạt động. Không cưỡng bức dừng giữa lúc ghi file.
- Export center hiển thị dung lượng, sheet, SHA-256, đối chiếu số dòng/tổng chi phí COPQ, bộ lọc B/C FTT, lot HFPA và số lượt QC/MA Re-Inspection. Sai lệch/kết quả dở dang cần xác nhận trước khi tải, kể cả ZIP. File bị thay đổi sau kiểm tra bị chặn tải. Lượt chạy cũ vẫn tải theo cơ chế trước.
- Tìm tool, đánh dấu yêu thích và lưu preset chế độ/nguồn/tháng trên trình duyệt. Preset không lưu file, dữ liệu ô hoặc thông tin đăng nhập. Danh sách file chỉ được quét khi chọn nguồn để trang mở nhanh.

`report_registry.json` chứa metadata, chế độ, nhóm đầu vào, schema, các bước và loại đầu ra. `report_contracts.json` chứa hợp đồng dữ liệu/giới hạn; adapter Python của tool mới có thể khai báo bằng `runner: "module:function"`, nhận `workspace`, `run`, `stage`. UI đọc registry, không cần tạo thêm thẻ/form thủ công. Quy tắc đối chiếu chuyên biệt vẫn cần được triển khai/test khi tích hợp một nghiệp vụ mới.

Server mặc định chỉ nghe `127.0.0.1`; nếu cần cách triển khai mạng nội bộ cũ, đặt `CHINGLUH_BIND_HOST` theo cấu hình của đơn vị. Luồng mới không gọi API AI hoặc gửi file ra cloud. Nhật ký kỹ thuật mới chỉ có bước, thời gian, loại lỗi và vị trí mã nguồn; không lưu giá trị ô/traceback thô. File nghiệp vụ vẫn nằm trong workspace/kết quả cục bộ.

Retention đề xuất 30 ngày cho các lượt mới đã kết thúc:

```powershell
python report_retention.py --days 30
# Chỉ sau khi kiểm tra danh sách: xóa cả workspace/đầu vào/kết quả hết hạn
python report_retention.py --days 30 --apply
```

Mặc định chỉ liệt kê. Không tự xóa lượt cũ, lượt đang chạy, dữ liệu Desktop hoặc file SFC trong `HFPAExports`.

### Kiểm thử

```powershell
python -m unittest discover -s tests -p "test_*.py"
node tests/test_sfc_dashboard.js
node tests/test_sfc_export_ui.js
node --check static/report_tools.js
# Trình duyệt Chrome riêng, chạy ẩn, dữ liệu giả lập:
$env:CHINGLUH_UI_TESTS='1'
python -m unittest discover -s tests -p "test_report_browser.py"
```

Không có bước build frontend: đây là Flask + HTML/JavaScript thuần. Các test engine chạy cả ba bộ xử lý thật trên dữ liệu giả lập, trong thư mục tạm, tạo Excel/PowerPoint/ZIP; không kết nối SFC. Cần có dependencies trong `requirements.txt` và Chrome cho test trình duyệt. Xem `docs/report_reliability.md` để biết phạm vi đối chiếu, API và phần chưa kiểm chứng trực tiếp.
