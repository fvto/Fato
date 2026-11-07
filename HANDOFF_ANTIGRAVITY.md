# Bàn giao cho Antigravity IDE

Cập nhật ngày 06/11/2026. Workspace: `C:\Users\User\Documents\Chingluh Dashboard`.

## Mục tiêu và yêu cầu đã chốt với người dùng

Tiếp tục **cập nhật code xuất báo cáo**, để người dùng xuất lại bằng Dashboard. Không chỉ sửa một file kết quả thủ công.

Luồng bắt buộc:

1. Tạo chart trong file Excel.
2. Hoàn thiện **toàn bộ design trong Excel**: màu, font, title, legend, nhãn, highlight, bố cục bên trong chart.
3. Dùng thao tác **Copy chart từ Excel rồi Paste sang PPT qua ứng dụng Office/WPS**.
4. Sau khi Paste, **trên PPT chỉ căn vị trí và kích thước chart, cùng vị trí/kích thước mũi tên**. Không redesign chart trên PPT.
5. Không chuyển chart thành ảnh (kể cả CopyPicture/PNG/EMF). Không quay lại tự chèn chart XML/workbook vào PPT. Người dùng đã thử Copy/Paste thủ công và thấy tránh được lỗi repair khi gửi người khác. Không Edit Data được trong PPT vẫn chấp nhận được, nhưng vẫn phải là đối tượng được Copy/Paste từ Excel.

Yêu cầu về hình thức vẫn phải giữ:

- Đúng design và bảng màu template quý.
- Các dòng `TOP ...` phải ở giữa chart, không đè lên thành phần khác.
- Tên model/lỗi dài phải đầy đủ, xuống dòng và dùng font nhỏ phù hợp; không tự bỏ nhãn ở giữa.
- Số liệu hiển thị phần trăm chỉ **một chữ số thập phân**; dữ liệu nguồn giữ độ chính xác.
- Trend có highlight như mẫu: VN xanh ngọc, Indo vàng, theo từng điểm số tổng vùng. Legend nhà máy phải hiển thị rõ trong WPS lẫn PowerPoint.
- Nhãn trend phải bố trí gọn, tránh đè đường và các nhãn khác.
- Mũi tên nằm ngay dưới cột model đầu tiên (Top 1), giữa chart model và chart lỗi.
- Các chart có vùng riêng, không chồng lấn nhau hay với các thành phần của template.

Các yêu cầu dữ liệu trước đó đã triển khai, cần giữ:

- Báo cáo quý FTT dùng cột **Model**, không dùng Model 2. Áp dụng cho **toàn bộ quý hiện tại**, bao gồm dữ liệu FTT đã cache trong Tableau. HFPA giữ logic đã được người dùng chấp nhận.
- Re-inspection: cùng một quốc gia phải cùng màu trong cả bốn biểu đồ tròn; OTHERS màu xám, không tô màu theo thứ hạng.

## Trạng thái thực tế — chưa hoàn tất luồng mới

Đã thay `editable_charts.convert()` bằng wrapper gọi adapter native Office mới. **Adapter hiện chưa chạy thành công với bước tách liên kết chart trong WPS. Không báo người dùng rằng đã sửa xong repair.**

Đã chạy được một lượt Copy/Paste native trước khi thêm tách liên kết:

- 12 chart được Copy/Paste thực sự bằng `ChartObject.Copy()` và `slide.Shapes.Paste()`.
- Office lưu PPTX, sau đó mở lại được; quá trình native có lúc chờ vài phút.
- Kiểm tra package chỉ đọc: slide 7 và slide 8 mỗi slide có **6 chart native**, không phải ảnh hay OLE trong lượt đó.
- Tuy nhiên cả 12 chart vẫn liên kết tới workbook tạm `file:///C:\Users\User\AppData\Local\Temp\office-qbr-...\source.xlsx`. Đây **không phải output có thể giao** vì thư mục tạm sẽ bị xóa.
- Output lần thử đó: `scratch/native-office-chart-check/HFPA-FTT-FY27-Q2.pptx` và `.xlsx`. Chỉ dùng chẩn đoán, không coi là bản cuối.

Đã thêm bước `shape.Chart.ChartData` / `data.IsLinked` / `data.BreakLink()` để dùng Office ngắt liên kết mà không đổi design. Lần thử sau lỗi COM:

```text
pywintypes.com_error:
(-2147352567, 'Exception occurred.',
 (0, None, None, None, 0, -2147467263), None)
```

`-2147467263` là E_NOTIMPL; cần xác định chính xác member của WPS trong đường gọi ChartData/IsLinked/BreakLink chưa được hỗ trợ. Lúc xuất traceback, file đang được chỉnh nên số dòng hiển thị lệch; đừng dựa vào dòng 180 trong log cũ để suy ra member.

Chưa chạy lại native export sau khi thêm `_check_links()`. Helper này chỉ đọc relationships, từ chối output còn trỏ về thư mục tạm. **Không sửa package PPT bằng XML để lách lỗi.**

## Code liên quan

### `report_tools/quarterly_quality/office_charts.py` — mới, chưa tracked

Adapter đang triển khai:

- Import pywin32 lazily; `CoInitialize`/`CoUninitialize`.
- Tạo Excel hoặc WPS (`Excel.Application`, `ket.Application`) và PowerPoint hoặc WPS (`PowerPoint.Application`, `kwpp.Application`).
- Làm trên bản sao Excel/PPT riêng trong thư mục tạm.
- Tìm chart Excel bằng worksheet và hàng nguồn trong series formula, không dựa vào số thứ tự chart.
- Chỉnh aspect ratio của chart **trong Excel trước Copy** cho trùng kích thước ô đặt trên PPT; tách vị trí chart trong Excel để tránh đè nhau.
- Dùng Copy/Paste native; chặn kết quả là picture (`Type` 11/13).
- Sau Paste chỉ đặt Name, Left, Top, Width, Height, LockAspectRatio, và căn mũi tên. Phần ChartData mới thêm chỉ phục vụ tách liên kết dữ liệu, không styling.
- `SaveCopyAs(...,24)` qua Office, kiểm tra không còn link tạm, đóng và mở lại PPT, rồi mới publish kết quả.
- `_slide()` đã xử lý việc WPS báo HasTextFrame nhưng không đọc Text được ở một số shape.
- Cleanup chỉ đóng document của chính job; chỉ Quit app khi collection document rỗng. **Không Quit bừa vì WPS có thể dùng chung instance với tài liệu người dùng.**

Hướng nghiên cứu tiếp: tìm cách Paste thành đối tượng tự chứa qua API Office/WPS. Có thể thử Paste Special **OLE object, không phải picture**, nếu phù hợp yêu cầu; không mặc định coi đó đã hoạt động. Default Paste hiện tạo chart linked và WPS thiếu một phần API ChartData. Dùng native writer để lưu kết quả, không tự viết chart XML vào PPT.

### `report_tools/quarterly_quality/editable_charts.py`

- Giữ toàn bộ thiết kế Excel trong `styled_chart()` và `prepare_excel_charts()`.
- `convert()` hiện chỉ gọi `office_charts.copy_charts()`. Code tự nhúng chart XML vào PPT trước đây đã bị gỡ khỏi hàm này.
- `CHART_SPECS`: định danh 12 chart bằng metric/region/kind/shape-id/sheet/source-rows.
- `chart_rectangle()`: 6 ô chart mỗi slide; trend bên trái, model và defect bên phải, có khoảng trống cho mũi tên.
- `RANK_PLOT_X=.04`, `RANK_PLOT_WIDTH=.92` dùng tính tâm cột Top 1.
- Nhãn dài dùng `strLit` với newline vì WPS bỏ qua newline trong cache của `strRef` và từng bỏ mất nhãn model ở giữa.
- Nhãn phần trăm có rich text tường minh `f'{value:.1%}'`; highlight theo từng điểm có cả fill và text highlight để WPS hiển thị.
- Legend dùng tên series literal, không formula chuỗi có dấu ngoặc kép, và không cố định manual legend bounds.
- Title auto layout, overlay false, paragraph centered. Font chart tối đa 10.5 pt; category font nhỏ.
- Nhãn trend xếp trên/giữa/dưới theo giá trị từng tháng. Cần xác minh native Excel save/Paste còn giữ design mong muốn bằng hình thực tế.

### `report_tools/quarterly_quality/template_output.py`

- `write_excel()` tạo cả 12 chart, kể cả vị trí ranking chưa có dữ liệu, rồi gọi `prepare_excel_charts()`.
- `write_powerpoint()` cập nhật header/date/consumer summary theo template. Nếu có workbook cùng tên, bỏ tạo ảnh chart rồi gọi `convert()`.
- Pipeline bình thường luôn xuất Excel trước PPT. Legacy image fallback khi không có workbook vẫn còn cho ngữ cảnh cũ/tests; tuyệt đối không dùng làm fallback cho lỗi native Copy/Paste của luồng bình thường.
- `defect_colors_for()` lấy màu theo tên lỗi từ Color_template, dùng màu thống nhất với Re-inspection khi cùng tên.

### Màu template hiện dùng

- Trend FTT VN: `002060`, `C9A779`, `8C9BAF`; highlight `66FFFF`.
- Trend FTT Indo: `6453AD`, `9A79C5`, `89CDC5`; highlight `FFFF66`.
- Trend HFPA VN: `4B375F`, `DBC9FF`, `7C5C78`; highlight `00FFFF`.
- Trend HFPA Indo: `4B375F`, `E0D4DF`, `DCF3FF`; highlight `FFD966`.
- Model FTT VN/Indo: `5B9BD5` / `9C7BC7`; HFPA VN/Indo: `5B9BD5` / `8C7CA2`.
- Cleanness `10253F`; Over cement `4F6228`; Thread end `4F7FBD`; Midsole/Outsole to upper bond gap `F79646`.

### Các file đã sửa ở các lượt trước

- `sfc_inputs.py`, `tableau_source.py`: FTT dùng Model; bookkeeping model coverage và dữ liệu theo ngày.
- `sfc_quarterly_quality_acquisition.py`, `sfc_report_acquisition.py`: lên lịch thay dữ liệu cache FTT chưa có Model coverage cho cả quý; giữ HFPA và các kỳ ngoài phạm vi, tránh cộng trùng.
- `report_tools/reinspection/pptx_generator.py`: màu quốc gia nhất quán, chuẩn hóa alias, màu unknown xác định ổn định.
- Các tests quarterly sources/tableau/charts/acquisition và `tests/test_reinspection_country_colors.py` bảo vệ những thay đổi trên.
- `requirements.txt`: vừa thêm `pywin32>=306; sys_platform == "win32"`.
- `report_runtime.py`: vừa thêm hiển thị `OfficeChartCopyError` bằng thông báo cụ thể.
- `README.md`: vừa cập nhật mô tả pipeline native; mô tả mục tiêu, **không phải bằng chứng luồng mới đã đạt**.

## Kiểm tra đã chạy

Lệnh gần nhất đã PASS:

```powershell
python -m unittest discover -s tests -p 'test_quarterly_quality*.py'
```

Kết quả: **35 tests pass**. Đây là tests logic dữ liệu và thiết kế Excel, **không chứng minh native export đã chạy xong hoặc không còn repair**.

`tests/test_quarterly_quality_presentation.py` vừa được chỉnh:

- Bỏ assertions dựa vào cấu trúc PPT do code tự nhúng chart XML cũ.
- Giữ checks thiết kế hoàn chỉnh trong Excel: format 0.0%, font, title center, tên dài đầy đủ, palette, series names/highlight.
- Giữ test wiring `write_powerpoint()` gọi converter và không tạo ảnh chart khi có peer workbook.
- Cần bổ sung tests adapter native: Copy trước Paste, không styling trên PPT, không chấp nhận picture, không link file tạm, không đóng tài liệu người dùng, lỗi không publish output sai.

Trước lượt native mới, tests Re-inspection country colors và acquisition cũng từng pass (14 tests). Không chạy lại các tests đó trong lượt cuối.

## Chạy lại native export và kiểm tra kết quả

```powershell
python scratch/verify_chart_pipeline.py scratch/native-office-final-check
```

Hoặc có dump stack định kỳ để tìm COM call chờ lâu:

```powershell
python scratch/verify_native_debug.py scratch/native-office-final-check
```

Builder scratch đọc dữ liệu từ:

`Report/ToolRuns/499b6f10f5b2410f95c437febfc5b111/workspace/Output/HFPA-FTT-FY27-Q2.xlsx`

Template:

- `Report/Quarterly/Partner QBR Format Q2 FY27-CLG.pptx` (8 slide, FTT slide 7, HFPA slide 8).
- Các path Excel template lấy từ `template_output.template_paths()`.
- Slide size: 960 × 540 pt, tương đương 12192000 × 6858000 EMU.

Tiêu chí hoàn thành:

1. Native Copy/Paste xuất được đủ 12 đối tượng chart tự chứa; không có đường dẫn Excel tạm trong file cuối.
2. Office/WPS tự lưu rồi mở lại được output; kiểm tra thêm PowerPoint nếu có. Không kết luận hết repair chỉ từ package validator.
3. So sánh trực quan slide 7/8 với mẫu: đầy đủ tên dài, title TOP centered, palette đúng, highlight rõ, legend đúng, nhãn không đè nhau, mũi tên thẳng cột đầu.
4. Kiểm tra design trước/sau native save. Chỉ sửa design ở Excel, PPT chỉ alignment.
5. Tests phù hợp pass; README mô tả đúng khả năng đã xác minh.
6. Báo lại người dùng code đã cập nhật để họ xuất lại; phân biệt kiểm tra trên máy này và việc gửi/mở trên máy khác chưa thử.

## Môi trường và bảo vệ dữ liệu người dùng

- Windows, PowerShell. `python` hiện là Python 3.13, đã có pywin32 và dependencies của project.
- COM đã thực sự khởi tạo được Excel và WPS Presentation trên máy này. `PowerPoint.Application` có lúc báo application busy; adapter retry rồi thử `kwpp.Application`.
- Runtime Python bundled của Codex thiếu một số dependency project/pywin32; dùng `python` của project cho native test.
- Dashboard ở `http://127.0.0.1:5050`, server Python PID gần nhất 15988. Không restart khi có job người dùng đang chạy. Worker mới thường import code mới.
- Git đang có nhiều thay đổi chưa commit. **Không reset/checkout/restore cả workspace.**
- Hai file `Report/Quarterly/Tableau data source 2026 HFPA Bonding.xlsx` và `.coverage.json` đã thay đổi trong các job dữ liệu của người dùng; không coi là file tạm do agent tạo, không revert/xóa.
- Không tự đóng toàn bộ WPS/Excel/PowerPoint, không kill app người dùng. Chỉ đóng tài liệu do lượt test của mình mở.
- File ở `scratch/chart-title-arrow-check/` và `scratch/chart-palette-check/` là kết quả cũ của luồng tự nhúng XML. Có thể tham khảo design, **không giao thay cho native output mới**.

## Việc tiếp theo nên làm ngay

Xác định cách WPS Paste chart tự chứa mà vẫn giữ đối tượng chart và design Excel, xử lý API chưa hỗ trợ; sau đó kiểm tra save/reopen và đối chiếu trực quan. Đây là việc còn dang dở quan trọng nhất. Không quay lại giải pháp ảnh hoặc tự nhúng chart XML chỉ để test pass.
