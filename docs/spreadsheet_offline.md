# So sánh Excel offline và xuất HTML / PowerPoint

Mở Dashboard ở `http://localhost:5050/compare`. Giữ server Dashboard chạy trên máy; chế độ **So sánh offline** không cần Internet, ChatGPT, Codex hay API AI.

## Kiểm tra nhanh số dòng và %

Nút **Kiểm tra số dòng và %** là lựa chọn bổ sung. Nút chính **Xem gợi ý so sánh** và luồng so sánh hiện có được giữ nguyên.

Chọn File A / File B rồi bấm **Kiểm tra số dòng và %**. Kết quả báo riêng số dòng có giống nhau và giá trị % có giống nhau, kèm từng ô khác hoặc chưa xác minh được.

- Đếm các dòng có nội dung, tính cả tiêu đề và dòng tổng, bỏ dòng trống và dòng chỉ có định dạng. Kiểm tra riêng từng trang tính; tổng số dòng bằng nhau không che giấu chênh lệch giữa các trang.
- Ghép trang tính cùng tên. Nếu mỗi file chỉ có một trang tính, ghép hai trang dù tên khác nhau. Trang còn thiếu được báo riêng.
- Chỉ đối chiếu ô có giá trị kèm `%` hoặc định dạng phần trăm, theo cùng vị trí dòng/cột. Không tự ghép theo mã, sắp xếp dòng/cột, tính lại tổng hoặc trung bình. Các cột văn bản và số thường không được so sánh trong nút này.
- So sánh chính xác, không làm tròn, độc lập với lựa chọn làm tròn của luồng chính. Nhận diện tương đương `95%` dạng văn bản và `0.95` có định dạng phần trăm. Số thường `95` không tự được hiểu thành `95%`.
- Không tìm thấy ô % sẽ báo **Không tìm thấy ô % để kiểm tra**. Công thức chưa có kết quả đã lưu, lỗi Excel và giá trị thiếu được báo cần xem lại, không coi là khớp.

Kết quả và tải CSV/JSON được lưu cục bộ; tải lại trang chỉ khôi phục kết quả, không tự chạy lại.

Chênh lệch số dòng hiển thị ngay ở đầu kết quả, ghi rõ lệch bao nhiêu dòng và file A hay B nhiều hơn. Bảng số dòng từng trang tính luôn mở, gồm số dòng hai bên và chênh lệch. Kết quả đã lưu từ bản trước cũng hiển thị được chênh lệch khi tải lại trang, không cần chạy lại.

## Workflow tự động

1. Thêm hai file Excel. File đầu tiên là **File A · dùng làm bản gốc**, file tiếp theo là **File B · cần kiểm tra**; có thể đổi lại trước khi chạy.
2. Bấm **Xem gợi ý so sánh**. Trang hiển thị cặp trang tính, cột ghép dòng, phạm vi và lưu ý. Bước này chưa chạy so sánh.
3. Bấm **So sánh theo gợi ý** để xác nhận và xem chênh lệch. Có thể xuất báo cáo từ kết quả.

Giao diện mặc định thu gọn file đã lưu, loại báo cáo, nhà máy và thiết lập thủ công. **Tùy chọn so sánh** cho phép đổi loại báo cáo hoặc làm tròn; số chữ số làm tròn luôn được hiển thị ở luồng chính. **Nâng cao · Tự chọn cột, vùng dữ liệu và phép tính** giữ toàn bộ công cụ thủ công và mẫu đã lưu. Khi không tự ghép được bảng, nút **Tự chọn dữ liệu cần so sánh** mở phần này với đúng hai file đang chọn.

Bấm **So sánh theo gợi ý** là bước xác nhận phạm vi/phương pháp; không cần đánh dấu thêm ô xác nhận. Khi thay file hoặc tùy chọn, gợi ý cũ bị hủy và cần xem lại. Tải lại trang khôi phục kết quả đã lưu, không tự chạy lại. Báo cáo giữ phương pháp, cảnh báo, phạm vi, tên nguồn/SHA-256, vị trí và ảnh đầu vào.

Workflow có hai mẫu chạy offline, được nhận diện lại từ tiêu đề mỗi lần nhập file:

- **MQAA:** bảng chi tiết có Audit Date / Audit form / Factory / Line/Machine / MQAA% và ma trận có Line / Average / GRAND TOTAL. Dùng sheet Combined để tránh cộng cả Combined lẫn các sheet nguồn. Tự kiểm tra Average từng dòng, Grand Total từng cột, đối chiếu tỷ lệ theo Line × Audit form trong đúng nhà máy, và phát hiện lỗi Excel/#REF!. Phạm vi nhà máy được lấy từ chi tiết QC hoặc phải chỉ định rõ nếu không có. Kỳ khác nhau hoặc vai trò file bị đảo sẽ bị chặn.
- **Bảng cùng cấu trúc:** ghép sheet theo tên và bộ tiêu đề; ghép cột theo nhãn kể cả khi đổi thứ tự. Đề xuất khóa duy nhất từ một đến ba cột mã/văn bản, nhưng người vận hành phải xác nhận ý nghĩa khóa. Sheet chưa ghép được được liệt kê, không bị báo là đã kiểm tra.

MQAA kiểm trung bình trên tỷ lệ đang hiển thị, bỏ trống/N/A và giữ 0. **Sai số bắt buộc bằng 0 sau bước làm tròn đã chọn.** Bộ chọn **Làm tròn trước khi so sánh** có các mức không làm tròn / 1 / 2 / 3 chữ số. Tính trung bình bằng dữ liệu đầy đủ, sau đó làm tròn kết quả và giá trị QC theo `ROUND_HALF_UP` rồi so sánh; không làm tròn từng đầu vào trước khi tính. Ví dụ `97.625%` thành `97.6%` khi chọn 1 chữ số. Kết quả giữ cả giá trị trước và sau làm tròn để truy vết. Kế hoạch phiên bản cũ phải nhận diện lại. Đối chiếu với nguồn dùng trung bình tỷ lệ theo dòng export, giữ nguyên dòng lặp; chưa có Audit ID thì không tự gộp những dòng nghi là cùng lượt audit.

Mặc định đã chọn **1 chữ số như QC** theo yêu cầu. Việc đổi mức làm tròn áp dụng cho lần chạy mới; kết quả đã lưu không tự tính lại. Báo cáo hiển thị tối đa 3 số lẻ và bỏ số 0 thừa: `97.600%` thành `97.6%`, `100.000%` thành `100%`.

Các mẫu có bố cục/ý nghĩa chưa nhận diện được cần thiết lập mapping thủ công và lưu preset. Không suy luận ý nghĩa nghiệp vụ từ vị trí ô hoặc dùng AI để tự chọn phương pháp. Bản sao đầu vào bị thay đổi sau khi lập kế hoạch sẽ bị chặn chạy; cần import lại để tạo kế hoạch mới.

## Thiết lập thủ công

1. Upload hai file Excel hoặc chọn từng export và bấm **Import**.
2. Chọn **So sánh offline**, bấm **Thiết lập so sánh 2 file**.
3. Chọn file A/B, sheet và vùng dữ liệu bỏ dòng tiêu đề, ví dụ `A2:D100`.
4. Chọn cách ghép dòng. Nếu dòng có thể đổi thứ tự, dùng cột khóa như `A,B` (Factory + Line). Nếu hai bảng có cùng bố cục, có thể ghép theo thứ tự dòng.
5. Nhập các cột giá trị của hai bên theo cùng thứ tự, ví dụ `C,D`. Chọn cách hiểu số/phần trăm và sai số cho phép.
6. Bấm **Chạy so sánh offline**. Kết quả chỉ rõ hai giá trị, chênh lệch và địa chỉ ô; dòng thiếu, khóa trùng và lỗi công thức được báo riêng.

Thêm quy tắc cho mỗi cặp sheet / phần dữ liệu. Có thể lưu preset để dùng lại thiết lập. Bộ so sánh không tự suy luận ý nghĩa các phần hoặc diễn giải yêu cầu bằng ngôn ngữ tự nhiên. Công thức được kiểm tra bằng kết quả đã lưu: tính lại và lưu trong Excel / WPS trước khi import.

## Gửi báo cáo cho người khác

Sau khi chạy, chọn **HTML** hoặc **PowerPoint (.pptx)** trong **Định dạng báo cáo**, rồi bấm **Xuất báo cáo**. Cả hai gom theo sheet QC, bên trong giữ đúng thứ tự dòng rồi cột của file QC, không sắp xếp theo tên chữ cái hoặc thứ tự tạo quy tắc.

HTML mặc định chỉ hiện bảng tổng hợp ngắn theo sheet; chi tiết, nguồn/phương pháp và ảnh nằm trong phần mở rộng. File tự chứa ảnh và mở offline được. PowerPoint có một trang tổng hợp cho case MQAA mẫu; phần chi tiết và ảnh gom trong phụ lục. Bảng và chữ là các đối tượng chỉnh sửa được. **Phần trăm hiển thị tối đa 3 số lẻ, bỏ số 0 thừa**, ví dụ `93.678%`, cho cả Dashboard / HTML / PPTX. Số hiển thị này độc lập với mức làm tròn dùng khi so sánh; báo cáo ghi rõ mức tính đã chọn và sai số 0. PowerPoint giữ giá trị gốc trong ghi chú của trang. Chênh lệch dùng nhãn `%` và phép trừ hai tỷ lệ, không đổi thành tỷ lệ tăng/giảm tương đối.

Ảnh được chụp bằng Chrome từ vùng dữ liệu Excel dựng lại trên máy; không giữ nguyên kiểu chữ, định dạng, biểu đồ hoặc bố cục trang in của Excel gốc. Ảnh PowerPoint gom vùng lớn để giảm số trang, tập trung các ô cần kiểm tra; HTML giữ thêm vùng dữ liệu nguồn để truy vết.

Vàng đánh dấu ô cần xem lại; xanh đánh dấu đầu vào để tính / đối chiếu, không kết luận các ô xanh sai. Báo cáo có bảng xử lý nhanh và liên kết đến từng ảnh / chi tiết. Lỗi phép tính nội bộ, chênh lệch với nguồn và lỗi công thức được phân nhóm riêng.

Gửi trực tiếp file HTML. Người nhận mở bằng trình duyệt, không cần Dashboard, Internet hoặc file Excel gốc. Ảnh được nhúng trong HTML; không có tài nguyên hoặc mã theo dõi bên ngoài. Báo cáo có thể chứa các ô lân cận để cung cấp ngữ cảnh, nên xem lại trước khi gửi.

Báo cáo bao gồm tất cả điểm cần xem lại của lần chạy, kể cả ngoài trang/bộ lọc đang xem. Chỉ giá trị khớp chính xác được loại khỏi danh sách chênh lệch. Một báo cáo hỗ trợ tối đa 100 vùng ảnh chi tiết; nếu vượt giới hạn, chia nhỏ quy tắc/vùng dữ liệu rồi xuất từng báo cáo. Không âm thầm bỏ ảnh hoặc bỏ vấn đề. Kết quả cũ chưa chạy chế độ sai số 0 cần chạy lại trước khi xuất.

## Xóa toàn bộ export

**Xóa tất cả file** hoạt động không cần chọn riêng một file. Hộp xác nhận ghi tổng số export; xóa cả các file vượt ngoài 80 mục mới nhất đang hiển thị. Chỉ xử lý file Excel thuộc nguồn danh sách export, không xóa thư mục, metadata hoặc bản sao đã import. Export mới xuất hiện sau hộp xác nhận không thuộc đợt xóa đó. Nếu file bị khóa, danh sách cập nhật và nêu file chưa xóa được để thử lại.
