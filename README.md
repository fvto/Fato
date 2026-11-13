# Chingluh Dashboard

Giao diện hiện tại dùng phong cách tối giản lấy cảm hứng từ Nothing trên cả trang chính và `/reports`: nền trắng ngà, tiêu đề serif, wordmark dạng chấm, menu nổi và chuyển động có hỗ trợ giảm motion. Nhấn **Ctrl+F5** để tải giao diện mới. Font/icon được phục vụ cục bộ. Xem `docs/nothing-workspace.md` cho danh sách thay đổi và kiểm thử.

Mở `MO_BANG_DIEU_KHIEN.bat`:

**QMS & PDF QA:** danh mục QMS được đồng bộ từ menu SFC hiện hành, gồm QA Manual và QA Document (`QADocument.aspx?id=QMS750`). Bấm **↓ PDF** cạnh hai mục này hoặc mở **QA Manual & QA Document · PDF Library**: tìm theo mã/tên giày, loại tài liệu, dự án; lọc nguồn, nhà máy và EN/VN rồi tải PDF gốc. Danh sách có phân trang và thời điểm cập nhật. Nút **Cập nhật từ SFC** đọc đủ mọi trang của cả hai nguồn cho nhà máy đã chọn; cần đăng nhập Chrome SFC. Danh mục cũ được giữ nếu quét chưa đủ hoặc lỗi kết nối. Nhà máy chưa có dữ liệu cần cập nhật lần đầu. PDF được kiểm tra trước khi gửi về dưới dạng file tải xuống; lỗi đăng nhập/trang xem không bị lưu nhầm thành PDF. Danh mục PDF nằm cục bộ trong `Report/QMSDocuments/catalog.json`; không lưu cookie vào danh mục. Có thể đồng bộ bằng `python sfc_memory_crawler.py --qms-only --plant VH`.

**Đối chiếu số lượng và tải hàng loạt:** nhóm **Bảng dữ liệu SFC** là mặc định; HFPA Guide và BDTS được tách riêng. Số dòng SFC và số liên kết PDF hiển thị riêng vì một dòng có thể có cả EN/VN. Tại VH, QA Manual có 421 dòng/421 liên kết PDF trong bảng và 25 PDF hướng dẫn bổ sung (18 HFPA, 7 BDTS); QA Document có 431 dòng/432 liên kết PDF. Một liên kết có thể trỏ tới file đã bị gỡ trên SFC. **Tải tất cả … PDF (ZIP)** lấy toàn bộ kết quả khớp nguồn/nhóm/ngôn ngữ/tìm kiếm trên mọi trang. Mỗi PDF thành công lưu riêng kèm SHA-256 trong thư mục lượt tải. Lỗi HTTP 404/410 được ghi nhận rồi tiếp tục các file phía sau; lỗi đăng nhập/kết nối/ghi file dừng lượt và giữ file đã lưu. Nếu thiếu file, ZIP mang tên **THIEU_n_PDF**, có danh sách **PDF_CHUA_TAI_DUOC.csv**, và giao diện ghi **CHƯA ĐỦ (x/y PDF)**. Bấm **Tiếp tục / thử lại file lỗi** để kiểm tra lại bản đã lưu và chỉ tải phần thiếu; không tự gửi lại file lỗi. Tải lại trang tiếp tục xem lượt cũ, không tạo lượt mới. ZIP/PDF/trạng thái lưu trong `Report/QMSDocuments/Downloads/`; giới hạn 5 GB PDF mỗi lượt. Lượt lỗi trước khi có cơ chế lưu riêng cần tải lại một lần.

Chọn riêng **QA Manual** hoặc **QA Document** trong **Phần tài liệu**. Nút tải, tiến độ, link và tên ZIP gắn với nguồn đã chọn; mỗi ZIP chỉ chứa PDF của nguồn đó theo bộ lọc hiện tại. Danh sách trộn hai nguồn bị chặn cả ở máy chủ. Khi chuyển nguồn, dashboard chỉ hiện lại lượt tải đã lưu đúng nguồn và nhà máy.

1. **Mở Dashboard**: chạy server và mở giao diện.
2. **Mở dashboard và DOM**: chạy server, mở kết nối SFC và giao diện. Đăng nhập SFC nếu được yêu cầu.

`MO_SERVER.bat` chạy riêng server tại http://localhost:5050/.

**Combine Excel:** mở `http://localhost:5050/combine` hoặc bấm **Combine Excel** trong Dashboard. Dùng chung cho bảng bán hàng, tồn kho, sản xuất, chất lượng và các bảng Excel có một dòng tiêu đề. Chọn/kéo thả hỗn hợp `.xls`, `.xlsx`, `.xlsm`, `.csv`, kiểm tra trang tính, dòng tiêu đề được gợi ý và dòng cuối rồi **Xem trước → Gộp & tạo file Excel → Tải file**. Có thể giữ mọi cột (ô thiếu để trống) hoặc chỉ gộp khi cùng bộ cột, kể cả khác thứ tự. Tên cột được chuẩn hóa Unicode, khoảng trắng và hoa thường; không tự ghép cột khác tên/đơn vị. Chỉ bỏ cột trống nếu dữ liệu bên dưới cũng trống. Sheet `Sources` ghi phạm vi, số dòng và cột thiếu của từng bảng. Mẫu xuất đặc thù được nhận diện theo cấu trúc trong `spreadsheet_combine_profiles.py`; dòng tổng chỉ bỏ sau đối chiếu, có thể chọn giữ. Với mẫu chưa nhận diện, dùng dòng cuối để bỏ phần tổng/ghi chú. Gộp lại bản đã xuất vẫn giữ cột nguồn cũ, thêm cột truy vết mới với tên riêng. Công thức dùng kết quả đã lưu; không chạy công thức hoặc macro. Giới hạn: **150 file/lượt**, 200 MB/file, tổng file tối đa 5 GB. Không còn chặn ở 500.000 ô; bảng lớn được đọc/ghi theo luồng và chuyển sang gói Power Query khi trên 1.000.000 dòng dữ liệu. CSV được đọc trực tiếp khi gộp, giữ nguyên file nguồn. CSV nhận diện UTF-8 (có/không BOM), UTF-16/32 có BOM và suy đoán encoding cũ; hỗ trợ dấu phẩy, chấm phẩy, Tab, gạch đứng và chỉ thị `sep=`. Có thể chọn encoding/delimiter trong **Tùy chọn CSV** trước khi tải; xem [chi tiết CSV và giới hạn](docs/spreadsheet-combine.md). Dữ liệu nằm trong `Report/CombinedExcel`. Bảng lớn dùng bộ xuất theo luồng tiết kiệm bộ nhớ; bảng nhỏ dùng runtime Artifact Tool của Codex, cấu hình đường dẫn qua `CHINGLUH_ARTIFACT_NODE` và `CHINGLUH_ARTIFACT_MODULES`.

**Tra cứu Combine:** sau khi xem trước, bấm **Chuẩn bị tra cứu** để tìm theo một hoặc nhiều cột, ví dụ Mold Code, Size hoặc FTY. Bảng hiển thị Grand Total Number of Pairs riêng theo Size và Size Range, cùng tối đa 100 dòng chi tiết mỗi trang. Mold Code bỏ qua khoảng trắng; Size Range giữ nguyên hậu tố vị trí; lọc chính xác 8A1A2 để loại 8A1A2-VQ. Top 3 Issue chỉ lấy Unit Type = DRR và xếp theo tổng Number of Pairs từng Size/Size Range. Số lượng không hợp lệ được cảnh báo. Dữ liệu tra cứu được lập chỉ mục cục bộ theo lượt chọn, cần thời gian và dung lượng đĩa khi dữ liệu rất lớn.

**Spreadsheet Chat:** mở `http://localhost:5050/compare`, tải file rồi gõ yêu cầu bằng tiếng Việt hoặc tiếng Anh. Excel MCP Server kết hợp Codex CLI hoặc Antigravity CLI tự đọc cấu trúc file, kiểm tra phép tính hoặc đối chiếu dữ liệu và trả bằng chứng theo ô; có thể hỏi tiếp trong cùng cuộc chat. Cài runtime bằng `python scripts/setup_excel_ai.py` và đăng nhập Codex CLI. Muốn dùng Antigravity, cài và đăng nhập `agy`, rồi chọn **Antigravity CLI** trong Cách phân tích. Nội dung liên quan được gửi đến AI qua tài khoản của chế độ đang chọn. Chế độ chọn quy tắc thủ công vẫn có trong Advanced. Xem `docs/spreadsheet-compare.md`.

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

Chart được chỉnh dữ liệu và định dạng trong Excel, sau đó Copy/Paste bằng Office sang PowerPoint; không sửa chart hoặc cache trong gói PPTX sau khi lưu. Note HFPA bên phải slide Consumer Promise Review là ảnh chụp vùng `A1:E4` của sheet `QBR HFPA Summary`: ba tháng của quý được chọn, dòng VN/Indo, cột Avg và trung bình hai vùng bên dưới. Tháng, số liệu và công thức thay đổi theo mỗi lần chạy; ảnh mẫu chỉ quy định bố cục. Tháng chưa có dữ liệu để trống và chưa hiển thị trung bình đủ quý. WPS liên kết chart với Excel kết quả đi kèm; giữ Excel ở vị trí xuất nếu cần tiếp tục chỉnh chart.

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
- Đối chiếu số lượt (số dòng dữ liệu) riêng từng nhà máy trong kỳ: QC ≥ số dòng HFPA có `Pass/Fail = Fail`, MA ≥ số dòng HFPA Fail và MA ≥ QC. Không cộng QC với MA, không dùng tổng số đôi hoặc số lỗi khuyết tật. Nếu không đạt, báo rõ số thực tế, số yêu cầu và số lượt thiếu; chặn tạo báo cáo cả khi chọn file thủ công lẫn tự tải SFC.
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
