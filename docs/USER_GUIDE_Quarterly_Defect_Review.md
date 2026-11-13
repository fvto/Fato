# Hướng dẫn sử dụng - Rà soát lỗi NIKE theo quý (`Quarterly NIKE Defect Review`)

## Kiểm tra file bằng WPS (12/11/2026)

- Sau khi lưu PPTX, tool đóng tài liệu Excel và PowerPoint, bỏ kết nối điều khiển cũ, rồi tạo kết nối WPS Presentation mới để mở lại file. Việc đóng WPS Spreadsheets có thể làm mất hiệu lực kết nối Presentation cũ.
- Kết nối bị ngắt khi đóng tài liệu đã lưu được xử lý riêng; lỗi quyền truy cập hoặc lỗi mở file vẫn làm báo cáo dừng. Kết nối mới chỉ được thử lại tối đa 3 lần với lỗi COM tạm thời.
- Chỉ xác nhận hoàn tất sau khi mở lại đủ slide, đủ biểu đồ native và xuất ảnh xem trước. File vẫn được lưu qua WPS/Office; không sửa XML trong PPTX sau khi lưu.
- Lỗi xuất có ghi rõ thao tác đang thực hiện trong `review_summary.json`.

## Cập nhật luồng tải nguồn NIKE qua DOM (12/11/2026)

- Nguồn tự động của riêng báo cáo này dùng Chrome SFC đang kết nối DOM cho cả HFPA và FTT. HFPA bật NIKE, xóa lọc chuyền, đặt ngày rồi Search/Excel; FTT mở Select Station và dùng nút Excel NIKE.
- HFPA NIKE chia tối đa 7 ngày mỗi lần xuất DOM (01-07, 08-14, 15-21, 22-28 và phần cuối tháng), kiểm tra rồi ghép thành một file tháng. Các phần hoàn tất được giữ để dùng lại khi chạy tiếp. FTT NIKE vẫn xuất một lần theo tháng. Với khoảng ngày người dùng chủ động chọn ngắn hơn tháng, chỉ chia trong đúng khoảng đã chọn.
- File tải xong phải đúng schema NIKE, nhà máy và khoảng ngày trước khi đưa vào workspace. Cache sai chế độ hoặc sai kỳ sẽ được yêu cầu tải lại.
- Khi trình duyệt báo bận hoặc timeout, tác vụ dừng để tránh gửi thêm lệnh xuất trong khi SFC còn xử lý. Các file hoàn tất được giữ lại. Hủy tác vụ chỉ có hiệu lực ở điểm an toàn.
- Danh mục chuyền Factory Sandals vẫn được đọc qua phiên HTTP đã đăng nhập; đây là bước lấy danh mục, không phải luồng xuất file HFPA/FTT. Các báo cáo khác giữ nguyên cách tải trước đó.

Phân hệ **Rà soát lỗi NIKE theo quý** (`quarterly_defect_review`) được xây dựng **tách biệt hoàn toàn** với các báo cáo Internal/Tableau hiện có (`quarterly_quality`, `hfpa`, `copq`, `reinspection`) để phân tích sâu cấu trúc lỗi cho **riêng nguồn NIKE**.

---

## 1. Nguyên tắc dữ liệu & ranh giới quan trọng

- **HFPA NIKE và FTT NIKE chạy độc lập 100%:** Hai chỉ số không bao giờ gộp chung tử số/mẫu số hay chèn chéo số liệu của nhau.
- **Chỉ dùng nguồn NIKE:** Hệ thống từ chối các file `Internal` (trả lỗi chặn xuất ngay ở bước kiểm tra đầu vào).
- **Kho lưu trữ riêng (`Cache/QuarterlyDefectReview`):** Không đọc, không ghi và không làm thay đổi file `Tableau data source 2026 HFPA Bonding.xlsx` hoặc cache báo cáo quý cũ.
- **Trạng thái DR (`UNCONFIRMED`):** Công thức DR (tên gọi, tử số, mẫu số) hiện ở trạng thái **chưa chốt chính thức**. Khi chốt sau này sẽ lấy trực tiếp từ nguồn HFPA NIKE và FTT NIKE đã chuẩn hóa; không tự đặt công thức và không lấy số từ ảnh mẫu Q2.

---

## 2. Cách chạy báo cáo trên giao diện `/reports`

1. Mở trang `/reports`, chọn công cụ **Rà soát lỗi NIKE theo quý**.
2. Chọn **Quý tài chính** (mặc định **`FY27-Q1`**: `2026-06-01` → `2026-08-31`) và chỉnh **Từ ngày / Đến ngày** nếu cần xem một phần quý hoặc ngày chốt riêng.
3. Chọn **Chế độ xử lý**:
   - **Excel và PowerPoint 7 slide (`all`)**: Tạo đủ workbook `.xlsx`, file trình chiếu `.pptx`, ảnh xem trước 7 slide và JSON phân tích.
   - **Chỉ xuất Excel (`excel`)**: Tạo workbook `.xlsx` và JSON phân tích nhanh.
   - **Chỉ kiểm tra dữ liệu (`check`)**: Kiểm tra file nguồn, mapping và đối chiếu control mà không xuất file.
4. Chọn **Nguồn dữ liệu**:
   - **Tự động (`auto`)**: Dùng dữ liệu cache riêng trong `Cache/QuarterlyDefectReview` hoặc tải phần thiếu từ SFC khi Chrome SFC đã đăng nhập (`HFPA` bật `ckbNIKE=True`; `FTT` nhấn `btnExcelNIKE` và lấy danh mục line `Factory Sandals`).
   - **Thủ công / File trong thư mục**: Tải lên 3 nhóm tách riêng:
     - `hfpa_nike`: File `.xls`/`.xlsx` HFPA NIKE.
     - `ftt_nike`: File `.xlsm`/`.xlsx` FTT NIKE (`FTT Template`).
     - `sandals_catalog` *(tùy chọn)*: File danh mục `line_catalog_factory_sandals.json` hoặc file FTT có cột `Factory Group = Factory Sandals`.

---

## 3. Quy tắc tách khu vực `CLG = VH + JV + VB`

Hệ thống phân vùng **loại trừ lẫn nhau** (không cộng trùng) và áp dụng **cùng một chuẩn phân loại** cho cả HFPA NIKE và FTT NIKE, trong khi vẫn giữ nguyên số liệu gốc của từng nguồn:

- **VB (Factory Sandals):** Các chuyền thuộc nhóm `Factory Sandals` (xác định từ `QA.aspx?tree_id=Q602` → `QADetail.aspx?factory=VT` hoặc từ `Factory Group = Factory Sandals` trong dữ liệu FTT).
- **VH:** Bao gồm `VH + VH2`, **trừ đi** các chuyền thuộc VB.
  - *Trường hợp đặc biệt:* Nếu dòng HFPA ghi `UccId = CL_VB` nhưng chuyền đó chưa xuất hiện trong danh mục Factory Sandals, hệ thống giữ nguyên ở **VH** và gắn cờ **`Chờ đối chiếu danh mục VB`** trên tab kiểm tra / Slide 7 để không phân bổ đoán mò.
- **JV:** Bao gồm `JV + JV2`.
- **CLG:** Tổng hợp chuẩn xác của `VH + JV + VB`.

---

## 4. Cách đọc 3 tầng phân tích, chiều Model & 2 mẫu số

1. **Tầng 1 — Top 3 nhóm lỗi CLG:** Xếp hạng theo **11 nhóm lỗi chuẩn HFPA** (`Wrinkle / Crease / Mis-shape`, `Bond gap / Bonding`, `Alignment L+R Symmetry`, `Contamination`, v.v.) trên phạm vi CLG, kèm bảng đóng góp % của `VH / JV / VB`.
2. **Tầng 2 — Top ≤3 lỗi chi tiết trong từng nhóm × phạm vi:**
   - Xếp hạng độc lập trong từng ô `(Nhóm, Phạm vi)` cho `CLG`, `VH`, `JV`, `VB` (top chi tiết của VH có thể khác top CLG).
   - **Chỉ hiển thị số lỗi thật sự quan sát được** (nếu chỉ có 1 hoặc 2 lỗi sẽ chỉ hiện 1 hoặc 2; không tự bịa lỗi thứ 3).
   - **Giới hạn độ sâu của dữ liệu SFC thực tế:**
     - **HFPA NIKE SFC:** Schema chuẩn chỉ có 11 cột nhóm (`tier1_group_only`), chưa có cột tên lỗi chi tiết cấp 3 — báo cáo sẽ ghi rõ tình trạng này thay vì dùng tỷ lệ FTT để chia giùm HFPA.
     - **FTT NIKE SFC:** Có 25 loại lỗi × 3 mức `Rework / B Grade / C Grade` (`tier2_subtype_only`).
3. **Chiều song song — Top ≤3 Model trong từng nhóm × phạm vi:**
   - Hiển thị song song **bên cạnh** (không bao giờ thay thế) bảng lỗi chi tiết.
   - Giữ nguyên định danh model gốc (không tự gộp `GS / PS / TD / WMNS`).
4. **Drill-down tách biệt 2 mẫu số:**
   - **Tỷ trọng trong nhóm tại phạm vi (%)** = `Số lỗi chi tiết tại khu / Tổng lỗi của nhóm tại khu đó`.
   - **Đóng góp khu vào cùng lỗi CLG (%)** = `Số lỗi chi tiết tại khu / Tổng lỗi của chính lỗi chi tiết đó trên toàn CLG`.

---

## 5. Kết quả đầu ra (Excel 5 sheet & PowerPoint 7 slide chuẩn `AGENTS.md`)

- **Workbook Excel (`Quarterly_NIKE_Defect_Review_<Kỳ>.xlsx`):**
  1. `Overview_and_Charts`: Bảng KPI `CLG/VH/JV/VB`, bảng **HFPA 3 tháng động (VN / Indo / CLG)** dùng công thức Excel (`SUMIFS`, `SUM`, `AVERAGE`) theo đúng 3 tháng của quý đang chọn, bảng Top 3 nhóm & Top ≤3 chi tiết, và 5 biểu đồ native.
  2. `Drilldown_Normalized`: Bảng dữ liệu chuẩn hóa làm nguồn công thức cho các bảng tổng hợp.
  3. `Detail_Breakdown_HFPA`: Phân tích 3 tầng và Top Model song song riêng cho HFPA NIKE.
  4. `Detail_Breakdown_FTT`: Phân tích 3 tầng, Top Model song song và vùng dữ liệu sắp xếp giảm dần (`M:N`) cho biểu đồ lỗi chi tiết FTT.
  5. `Catalog_Mapping_&_Review`: Danh mục chuyền `Factory Sandals (VB)`, các dòng `Chưa ánh xạ`, cờ `Chờ đối chiếu danh mục VB`, trạng thái DR và cấu hình phiên bản mapping.
- **PowerPoint (`Quarterly_NIKE_Defect_Review_<Kỳ>.pptx`):**
  - Toàn bộ biểu đồ được định dạng hoàn chỉnh trong Excel rồi Copy/Paste native sang PowerPoint qua ứng dụng Office/WPS; bảng ghi chú HFPA 3 tháng bên phải Slide 1 được chụp trực tiếp từ vùng bảng công thức động trong Excel.
  - Tự động đóng Excel trước khi mở lại `.pptx` để xác minh liên kết ngoài chỉ trỏ về đúng workbook `.xlsx` đích (không trỏ file tạm) và xuất 7 ảnh xem trước `Output/previews/slide_1.png` .. `slide_7.png`.

---

## 6. Xử lý nhãn `Chưa ánh xạ` & lưu quy tắc Mapping Override

- Nhãn lỗi mới không khớp từ điển chuẩn sẽ **không** bị đẩy ngầm vào nhóm `Other / Miscellaneous` mà được giữ ở trạng thái `Chưa ánh xạ`.
- Ngay trong panel kết quả của lần chạy trên `/reports`, mở mục **Kiểm tra mapping & cảnh báo**, chọn nhóm chuẩn HFPA phù hợp cho nhãn chưa ánh xạ và bấm **Lưu quy tắc** (`POST /api/report-tools/quarterly_defect_review/overrides`). Quy tắc được lưu tách riêng tại `Cache/QuarterlyDefectReview/mapping_overrides.json` và tự áp dụng cho lần chạy lại báo cáo tiếp theo.

