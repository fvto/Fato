# Combine Excel / CSV

Mở `/combine`, chọn hoặc kéo thả hỗn hợp `.xls`, `.xlsx`, `.xlsm`, `.csv`. Tối đa **150 file/lượt**, **200 MB/file**, tổng file **tối đa 5 GB**, **không áp trần 500.000 ô**. Các giới hạn dung lượng vẫn áp dụng dù chưa đủ 150 file. CSV được đọc trực tiếp và giữ lại hash file nguồn. CSV được xem là một bảng nguồn, kể cả khi vượt giới hạn dòng một sheet Excel. Số dòng nguồn vẫn truy ngược về bản ghi CSV gốc.

## Quy tắc gộp

Giữ nguyên bước làm sạch và ghép bảng hiện có: chọn dòng tiêu đề và dòng cuối, ghép theo tên cột đã chuẩn hóa Unicode/khoảng trắng/hoa thường, giữ hợp tất cả cột hoặc yêu cầu cùng bộ cột, bỏ dòng trống và tiêu đề lặp chính xác. Chỉ bỏ footer của mẫu đã nhận diện sau đối chiếu; có thể chọn giữ. Không suy đoán các cột khác tên là cùng dữ liệu. Không loại các dòng dữ liệu trùng nhau.

Kết quả vẫn là `.xlsx` gồm `Combined` và `Sources`, cùng ba cột truy vết nguồn. Các lượt gộp cũ và file đã xuất vẫn sử dụng được.

## CSV

- UTF-8 có/không BOM; UTF-16/32 có BOM được nhận diện trực tiếp. Toàn bộ file được kiểm tra encoding, tránh trường hợp đoạn đầu chỉ chứa ASCII nhưng có byte Windows/NBSP về sau. Các encoding cũ được suy đoán bằng `charset-normalizer`, có cảnh báo kiểm tra lại ký tự.
- Khi ký tự chưa đúng, mở **Tùy chọn CSV** trước khi tải lại để chọn UTF-8, UTF-16/32, Windows-1258/1252/1250, Latin-1, Shift-JIS, GB18030 hoặc Big5. Lựa chọn này áp dụng cho mọi CSV của lượt tải tiếp theo; Excel không bị ảnh hưởng. Chế độ tự động nhận diện riêng từng file.
- Dấu phân cách: phẩy, chấm phẩy, Tab hoặc gạch đứng; hỗ trợ chỉ thị đầu file `sep=` và lựa chọn thủ công. Dấu nháy kép, dấu phân cách trong ô và ô nhiều dòng được đọc theo quy tắc CSV.
- CSV giữ giá trị dạng chữ, gồm mã `001`, số dùng dấu phẩy thập phân, ngày tháng và chuỗi bắt đầu bằng `=` hoặc `'`. Quy tắc chuyển tỷ lệ phần trăm theo tên cột của chức năng cũ tiếp tục áp dụng. Không tự suy đoán kiểu số hoặc ngày tháng.
- Dòng nguồn CSV đếm theo **bản ghi**, kể cả khi một ô chứa xuống dòng. Dòng `sep=` vẫn chiếm vị trí đầu tiên, nên tiêu đề thường ở dòng 2.
- Encoding cũ của file ngắn và dấu phân cách trong nội dung có thể mơ hồ. Luôn kiểm tra bản xem trước; chọn lại encoding/delimiter và tải lại khi cần. Không dùng giải mã thay thế làm mất ký tự.

## Lỗi và bộ nhớ

Lỗi tải lên được liệt kê theo file; nếu có lỗi, toàn bộ lượt tải bị hủy và dọn thư mục tạm, không âm thầm bỏ nguồn. Lỗi chọn tiêu đề/phạm vi hoặc đọc lại nguồn ghi rõ tên file. Giữ kiểm tra file trùng nội dung, nguồn thay đổi và công thức Excel chưa có kết quả đã lưu.

Upload và hash đọc theo khối, CSV đọc từng bản ghi với mẫu nhận diện giới hạn 128 KiB; không giữ nội dung nhị phân của cả batch. Dữ liệu thô từng file được giải phóng trước file tiếp theo. Kiểm tra kích thước bảng hợp cột trước khi cấp phát kết quả, và kiểm tra XLSX sau xuất theo từng dòng. CSV giới hạn số cột và độ dài ô để phù hợp đầu ra Excel; bản ghi trống được xử lý theo luồng. Giới hạn 500.000 ô chỉ là ngưỡng chuyển sang chế độ xử lý lớn, không chặn upload hay tổng kết quả. Excel nguồn hỗ trợ tối đa 8 GiB sau giải nén mỗi file. Trên 1.000.000 dòng dữ liệu sau làm sạch: không xuất dữ liệu chi tiết thành XLSX; chuyển sang gói Power Query ZIP.

Bảng nhỏ dùng Artifact Tool; bảng trung bình có thể dùng openpyxl write-only. Nguồn lớn hoặc bảng hợp cột vượt ngưỡng RAM dùng luồng XML/ZIP64, giữ định dạng số/ngày/tỷ lệ, tiêu đề, bộ lọc, cố định dòng đầu và Sources. Thời gian phụ thuộc số dòng/cột và dung lượng giải nén; giao diện theo dõi tiến độ, không giữ tất cả dòng trong RAM.

## Kiểm thử

```powershell
$env:CHINGLUH_UI_TESTS='1'
python -m unittest discover -s tests -p 'test_spreadsheet_combine*.py' -v
```

Bao gồm các kiểm thử Excel cũ, CSV BOM/encoding/delimiter/nháy/xuống dòng, header và schema, lỗi theo file, trùng nội dung, giới hạn ô, 150 file hỗn hợp thật qua HTTP, từ chối file thứ 151, xuất XLSX và đọc lại, giao diện Chrome trên desktop/mobile, cùng kiểm thử lại bán hàng/tồn kho/lương và gộp lại đầu ra cũ. Chrome và runtime Artifact Tool cần có để chạy toàn bộ các ca giao diện/xuất thật.

Fixture `tests/fixtures/combine_legacy.xls` là workbook BIFF8 tổng hợp (ID/Count, 003/3), tạo bằng xlwt 1.3.0; chạy kiểm thử chỉ cần xlrd đã có trong ứng dụng.


## Luồng dữ liệu lớn và tiến độ

Upload trên 8 MiB được lưu vào thư mục tạm và xử lý nền với trạng thái tiến độ qua `/api/combine/tasks/<id>`. CSV được đọc trực tiếp theo luồng, không tạo XLSX trung gian hoặc nạp toàn bộ bảng vào RAM. Bảng lớn dùng cùng quy tắc header, schema, dòng trống, tiêu đề lặp, tỷ lệ, footer và truy vết của luồng cũ. Mặc định chọn sheet có nhiều dòng nhất trong workbook lớn; người dùng có thể đổi lựa chọn.

Xem trước chỉ giữ tối đa 12 dòng; thống kê đã xác minh được cache theo lựa chọn và hash nguồn. CSV tính sẵn thống kê làm sạch trong lần đọc khi tải lên, dùng cùng bộ lọc dòng với lúc gộp. Xem trước lớn lần đầu dùng lại thống kê này khi giữ đúng header và toàn bộ phạm vi nguồn; đổi header/phạm vi sẽ đọc lại, còn tùy chọn bỏ/giữ footer vẫn được đối chiếu riêng. Kiểm tra độ dài ô và ký tự điều khiển CSV theo nhóm tối đa 64 ô để giảm xử lý lặp và giữ bộ nhớ có giới hạn. Trước khi dùng cache hoặc xuất, kiểm tra lại hash file nguồn. Các lượt đã tải trước bản tối ưu vẫn hoạt động theo bộ nhớ đệm cũ. Xuất lớn được kiểm tra lại số dòng và hash giá trị từng sheet. Một lượt xử lý nền lớn chạy tại một thời điểm. Ứng dụng cần tiếp tục chạy trong quá trình đọc/gộp; giao diện hiện file và số dòng đang xử lý.


## Tra cứu và Grand Total theo size

Sau Xem trước, chọn **Chuẩn bị tra cứu**. Ứng dụng lập chỉ mục SQLite cục bộ theo các file/sheet và tùy chọn đang chọn; nguồn được kiểm tra hash trước khi công bố. Lượt chuẩn bị giống hệt được sử dụng lại. Chọn cột, cách khớp và nhập giá trị; các điều kiện áp dụng đồng thời. Khớp Mold Code bỏ khoảng trắng và không phân biệt hoa thường; chọn Có chứa cho tìm một phần.

Grand Total là tổng chính xác theo Decimal của cột Number of Pairs, tính riêng từng Size và Size Range. Size số như 9 và 9.0 được gộp; Size Range giữ nguyên hậu tố vị trí để phân biệt 8A1A2 và 8A1A2-VQ. Số không hợp lệ được đếm và cảnh báo, không âm thầm cộng. Chỉ hiển thị 1.000 nhóm đầu nếu có quá nhiều; bộ lọc giúp thu hẹp. Bảng chi tiết phân trang 100 dòng. Chỉ mục mã khuôn và size hỗ trợ khớp chính xác nhanh; lọc cột khác hoặc Có chứa có thể phải quét dữ liệu và mất thời gian. Chỉ mục cần thêm dung lượng ổ đĩa, đặc biệt với hàng chục triệu dòng.


## Ngưỡng xuất và Power Query

Upload trả về số dòng ước tính của các bảng được chọn mặc định. Sau Xem trước, dùng số dòng thực sau quy tắc làm sạch để quyết định: dưới hoặc bằng 1.000.000 dòng xuất XLSX; trên ngưỡng chỉ xuất gói Power Query (.zip), kể cả gọi API run trực tiếp. Gói gồm Data_*.csv UTF-8 BOM (mỗi phần tối đa 1 triệu dòng), Combine.pq, manifest số dòng/hash từng phần và Huong_dan.xlsx với ô cấu hình SourceFolder. CSV xuất giữ chuỗi mã, ký tự xuống dòng và cột truy vết. Các phần được kiểm tra CRC và số dòng trước công bố.

Máy hiện dùng WPS không hỗ trợ Queries qua API Excel. Vì vậy Huong_dan.xlsx là file hướng dẫn/cấu hình, không giả lập workbook đã nhúng Power Query. Người dùng Microsoft Excel có Power Query nhập Combine.pq qua Advanced Editor và chọn Only Create Connection hoặc Data Model, không tải toàn bộ dữ liệu vào worksheet. Truy vấn M chưa được thực thi trong Microsoft Excel trên máy này.

## Top 3 Issue DRR

Áp dụng toàn bộ điều kiện lọc hiện tại, sau đó chỉ lấy Unit Type đúng bằng DRR (bỏ khoảng trắng ngoài và không phân biệt hoa thường; NON-DRR không được lấy). Xếp Issue theo tổng Number of Pairs giảm dần trong mỗi Size và Size Range. Khi đồng hạng, sắp theo tên Issue và lấy đúng ba mục. Issue trống và số đôi không hợp lệ không tham gia xếp hạng; số lượng không hợp lệ vẫn được cảnh báo. Tối đa 1.000 nhóm được hiển thị; cần thu hẹp bộ lọc nếu nhiều hơn.
