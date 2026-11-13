# Review plan: phân tích lỗi quý HFPA, FTT và DR

Ngày lập: 11/11/2026. Trạng thái: kế hoạch để review và tích hợp, chưa thay đổi logic chạy báo cáo.

## 1. Yêu cầu đã được chốt

Đây là một báo cáo phân tích riêng, được truy cập từ tool báo cáo quý. Hai nguồn dữ liệu là HFPA NIKE và FTT NIKE. DR được lấy từ hai nguồn này theo xác nhận của người dùng, không yêu cầu file DR riêng và không mặc định DR là hàng trả về. Phân tích HFPA và FTT riêng; phần DR giữ dấu vết về nguồn HFPA/FTT tương ứng. Với mỗi nguồn, chọn ba nhóm lỗi lớn nhất của CLG theo số lượng lỗi; trong từng nhóm lỗi, chọn ba model có số lượng lỗi lớn nhất ở CLG, VH, JV và VB. Tính đóng góp của từng khu vào tổng số lượng của cùng nhóm lỗi tại CLG.

Tên viết đầy đủ và công thức DR chưa được người dùng xác nhận. Trước khi triển khai phần tỷ lệ DR, cần chốt tử số, mẫu số và cách trình bày theo từng nguồn. Không tự cộng HFPA với FTT để tạo một tổng DR.

**Chỉ báo cáo này sử dụng file HFPA NIKE và FTT NIKE. Không đổi các luồng báo cáo hiện tại sang NIKE.** Không thay nguồn của báo cáo HFPA tháng, FTT/COPQ, báo cáo quý đang có hoặc các tác vụ QAStation/MES410 khác. Không cập nhật nguồn Tableau dùng chung từ dữ liệu của báo cáo mới.

| Khu hiển thị | Phạm vi |
| --- | --- |
| VH | VH + VH2, sau khi loại toàn bộ chuyền sandals đã được xác định từ Group Factory Sandals ở FTT |
| JV | JV + JV2 |
| VB | Toàn bộ chuyền thuộc Group Factory Sandals ở FTT. HFPA không hiển thị Group; dùng danh sách chuyền đã được group sẵn tại FTT để nhận diện VB trong HFPA |
| CLG | VH + JV + VB, mỗi bản ghi chỉ thuộc một khu |

Q1 đang được hiểu là FY27 Q1: 01/06/2026–31/08/2026, phù hợp lịch quý trong template hiện tại. Đây là giả định đã nêu để người dùng có thể sửa; chọn quý trên giao diện phải luôn hiển thị khoảng ngày tương ứng.

Ảnh thứ nhất là tham chiếu bố cục. Số, quý Q2, tên model, tỷ lệ DR và các lỗi trong ảnh không phải số liệu đầu vào của Q1. Ảnh thứ hai là bảng quy đổi dạng lỗi NIKE do người dùng cung cấp.

## 2. Cách đọc kết quả mong muốn

Mỗi phần phân tích từ HFPA NIKE và FTT NIKE trả lời lần lượt; phần DR sử dụng lại kết quả có nguồn tương ứng:

1. Ba nhóm lỗi lớn nhất của CLG trong kỳ là gì, số lượng bao nhiêu, chiếm bao nhiêu trong tổng lỗi của phần đó?
2. Trong mỗi nhóm lỗi đó, VH, JV và VB đóng góp bao nhiêu lỗi và bao nhiêu phần trăm?
3. Với cùng nhóm lỗi, ba model nhiều lỗi nhất của CLG là gì? Ba model nhiều lỗi nhất trong từng khu là gì?
4. Những model và khu nào cần ưu tiên xem xét dựa trên số liệu trên?

Ví dụ minh họa, không phải kết quả Q1: CLG có 100 lỗi Bonding, gồm VH 40, JV 35 và VB 25. Đóng góp hiển thị lần lượt 40%, 35% và 25%. Không lấy 40 chia tổng mọi lỗi của VH để tính cột đóng góp này.

## 3. Hiện trạng đã kiểm tra và giới hạn dữ liệu

Các kết luận dưới đây dựa trên đọc mã nguồn và kiểm tra cấu trúc file, chưa chạy báo cáo NIKE hoặc kiểm tra bản xuất bằng Excel/PowerPoint/WPS.

| Hạng mục | Bằng chứng hiện tại | Hệ quả cho báo cáo mới |
| --- | --- | --- |
| Phân khu | `quarterly_quality.py` có VH, VH2, JV, JV2; tổng hợp VN, Indo, CLG | Cần chiều khu VH/JV/VB riêng, không sửa ý nghĩa các khóa cũ |
| Chọn lỗi | `sfc_inputs.py:load_sfc_quarter` chọn Top model rồi lấy lỗi của model đứng đầu | Cần engine lỗi trước, model sau; không tái sử dụng kết quả ranking cũ |
| Nguồn quý hiện tại | `sfc_quarterly_quality_acquisition.py` đang yêu cầu HFPA nội bộ / QAStation và FTT Internal | Giữ nguyên; viết acquisition riêng cho báo cáo NIKE này |
| Tableau | `Report/Quarterly/Tableau data source 2026 HFPA Bonding.xlsx`, sheet VH FTT có 9 cột: Date, Factory, Line1, Model Name, Inspection Qty1, Defect Qty1, Station1, Defect Issues1, Issues Qty1 | Không có cột Group sandals. Không đủ để tự chứng minh phân khu VB hoặc chế độ NIKE |
| File Q1 tìm thấy | Các file FTT Internal trong run `211e95b68b3e47e79fc0fe0cde0d1559`: tháng 6, tháng 7 và 01–03/08/2026; có bản đầu tháng 8 khác trong các run khác | Đây là nguồn nội bộ, không được đưa vào báo cáo mới. Tên file đầu tháng 8 cũng không chứng minh đủ cả Q1 |
| Xuất HFPA NIKE | `sfc_hfpa.py` và `sfc_excel.py` đã có tùy chọn NIKE và kiểm tra dạng export | Có thể dùng lại năng lực xuất với tham số riêng; không đổi mặc định dùng chung |
| Nguồn NIKE và phần DR | Chưa xác nhận được file mẫu NIKE đủ cấu trúc, dữ liệu và kỳ cho yêu cầu này. Người dùng xác nhận DR lấy từ HFPA NIKE và FTT NIKE | Xác minh hai mẫu NIKE; không tạo adapter hay yêu cầu file DR riêng |

Chưa có căn cứ để công bố Top 3 và phần trăm Q1 theo yêu cầu mới. Việc này cần nguồn HFPA NIKE, FTT NIKE đầy đủ kỳ và danh sách chuyền Group Factory Sandals. Phần DR được lấy từ hai nguồn đó. Không lấy số nội bộ, số từ ảnh hoặc bản Q2 thay thế.

## 4. Lấy dữ liệu và tách riêng chế độ báo cáo

Tạo loại báo cáo có mã riêng, ví dụ `quarterly_defect_review`, với tên giao diện “Phân tích lỗi quý HFPA / FTT / DR”. Có thể đặt cạnh lựa chọn báo cáo quý hiện tại.

| Đầu vào | Yêu cầu của nhánh mới |
| --- | --- |
| HFPA | Bản NIKE, đúng nhà máy và khoảng ngày. Kiểm tra chế độ trước và sau xuất |
| FTT | Bản NIKE, đúng kỳ và trạm/phạm vi chất lượng đã xác nhận từ màn hình FTT; không lấy FTT Internal |
| Danh sách sandals | Lấy danh sách chuyền từ đúng Group “Factory Sandals” ở mục FTT. Group trên giao diện không mặc nhiên là cột `FactoryGroup` trong file |
| DR | Kết quả từ HFPA NIKE và FTT NIKE; không phải đầu vào/file thứ ba. Công thức tỷ lệ cần được xác nhận trước khi triển khai |

Các tham số nguồn phải thuộc cấu hình nhánh mới. Không sửa `nike=False`, `PLANTS`, `REGIONS`, quy tắc hòa giải hoặc nguồn mặc định của nhánh cũ.

Mỗi lần xuất lưu thông tin: loại báo cáo, chế độ NIKE, nhà máy, Group được chọn, trạm/phạm vi lọc, từ ngày–đến ngày, thời điểm xuất, mã băm file và phiên bản danh sách chuyền. Khóa cache phải bao gồm các tham số này. File có sẵn đúng ngày nhưng sai chế độ không được tái sử dụng.

Không đồng thời cộng toàn bộ FTT với một bản xuất Group Factory Sandals: bản Group dùng để xác định thành viên hoặc đối chiếu; nếu làm nguồn số lượng phải phân hoạch rõ phần đã nằm trong tổng.

Nếu file NIKE thiếu mức chi tiết cần thiết, báo rõ trường thiếu và phần phân tích bị ảnh hưởng. Không âm thầm chuyển sang file nội bộ.

## 5. Dò chuyền và phân VH/JV/VB

**FTT là nguồn danh mục Group; HFPA là nguồn số liệu HFPA.** HFPA không hiển thị Group, đây là đặc điểm đã được người dùng xác nhận, không phải lỗi thiếu cột. Dò danh sách chuyền Factory Sandals từ FTT sang HFPA chỉ để phân khu; số lượng lỗi HFPA vẫn lấy từ chính HFPA NIKE, không lấy số lượng FTT gán sang.

Tạo bảng danh mục có tối thiểu: nhà máy gốc, xưởng, chuyền, Group gốc, khu báo cáo, hiệu lực từ ngày, hiệu lực đến ngày, nguồn chứng minh. Với chuyền ổn định cả quý có thể dùng một khoảng hiệu lực cho quý; nếu thay đổi Group phải lưu các khoảng tương ứng.

Khóa dò gồm nhà máy + xưởng + chuyền + thời điểm hiệu lực. Chuẩn hóa khoảng trắng, chữ hoa/thường và cách viết `Line1` / `Line 1`. Không bỏ nhà máy/xưởng rồi chỉ dò số chuyền. Không tự gộp hậu tố hoặc các mã chuyền khác nhau.

Quy trình:

1. Đọc toàn bộ danh sách chuyền của Group Factory Sandals ở FTT trong phạm vi báo cáo.
2. Lưu bảng đối chiếu cả mã gốc và mã chuẩn hóa, không cố định chuyền theo tên model có chữ Slide/Sandal.
3. Nhận diện VB trong FTT từ Group Factory Sandals và lưu danh sách chuyền tương ứng. Dò từng bản ghi HFPA vào danh sách này dù HFPA không có cột Group. Chuyền sandals thuộc VH/VH2 được tính vào VB ở cả hai nguồn, đồng thời loại khỏi VH trong chính từng nguồn.
4. Các bản ghi VH/VH2 còn lại vào VH. JV/JV2 vào JV theo quy tắc đã chốt.
5. Nếu danh sách Factory Sandals có chuyền thuộc JV/JV2, dừng riêng trường hợp xung đột này để xác minh vì có thể chồng quy tắc JV và VB. Không cộng một bản ghi hai lần.
6. Xuất bảng chuyền chưa khớp hoặc khớp nhiều khu. Không mặc định những dòng chưa xác định sandals là VH.

Phần DR kế thừa khu đã xác định từ các bản ghi HFPA NIKE/FTT NIKE tương ứng. Không cần một danh mục khu DR hoặc nguồn trả hàng riêng. Chỉ có tên model là chưa đủ dò chuyền. Không phân bổ số liệu HFPA theo tỷ trọng FTT hoặc ngược lại để ước tính số liệu thực tế.

Ví dụ kiểm thử giả định: Group Factory Sandals của FTT chứa khóa `VH / N14 / Line 1`. Bản ghi HFPA NIKE có đúng khóa này được tính vào VB dù không có Group. Một bản ghi `JV / N14 / Line 1` không tự động khớp vì khác nhà máy. Ví dụ này không khẳng định đó là danh sách chuyền thực tế.

## 6. Bảng quy đổi lỗi NIKE

Dùng nhóm HFPA ở cột trái ảnh làm nhóm chuẩn chung. Nhãn FTT ở cột giữa được quy về nhóm đó. Giữ nguyên nhãn gốc và thêm nhãn chuẩn, không ghi đè dữ liệu nguồn.

| Nhóm chuẩn theo ảnh HFPA | Các nhãn FTT tương ứng trong ảnh NIKE |
| --- | --- |
| Airbag | Airbag defect |
| Alignment L+R Symmetry | Alignment L+R Symmetry |
| Bond gap / Bonding | Bond gap / Rat hole; Delamination |
| Cement | Over cement |
| Contamination | Contamination |
| Interior (including sock-liner) | Interior defect; Accessories defect |
| Poor Quality Spray / Paint | Color / Paint migration, bleeding; Color mis-match; Paint surface quality |
| Scratch / Tear / Rip / High Buffing | Material damaged; Holes quality; Over buffing |
| Stitching / Loose Thread | Jump / Broken / Loose Stitch; Thread end; Stitching margin |
| Wrinkle / Crease / Mis-shape | Off-center; Rocking; Toe spring; Wrinkle or Deformed Bottom; Wrinkle or Deformed Upper; X-ray |
| Other / Miscellaneous | Yellowing; Other defects |

Đây là bản chép bảng từ ảnh để triển khai; cần đối chiếu cách viết thực tế của nhãn khi đọc file NIKE. `X-ray` được đặt theo ô nhóm Wrinkle / Crease / Mis-shape trong ảnh. Không tự thêm nhóm Outsole chỉ vì xuất hiện ở ảnh báo cáo DR; bảng quy đổi lỗi do người dùng cung cấp chưa có nhóm đó.

Mỗi mapping có nguồn áp dụng (HFPA/FTT), mã hoặc nhãn gốc, nhóm chuẩn, phiên bản và căn cứ. Phần DR kế thừa mapping từ nguồn tương ứng. Ưu tiên mã lỗi nếu file có mã. Các alias như Cleanness → Contamination chỉ được bổ sung khi nội dung mô tả thực tế xác nhận; không tự quy mọi lỗi “bẩn” hoặc mọi nhãn chứa từ tương tự vào một nhóm.

Không tạo một bộ mã DR riêng. Các nhãn HFPA/FTT chưa xác định vào danh sách “Chưa ánh xạ” có số lượng và nguồn; không tự cho vào Other. Nhóm Other chính thức vẫn được xếp hạng bình thường nếu đứng trong Top 3.

## 7. Đơn vị tính và chống cộng trùng

Tách hai lớp dữ liệu:

- Lớp kiểm tra/sản xuất: số đôi kiểm, sản xuất và lỗi ở đúng mức bản ghi nguồn NIKE.
- Lớp lỗi: một nhóm lỗi, model, khu và số lượng lỗi từ chi tiết. Dùng lớp này để chọn Top 3 và tính đóng góp khu.

Một lần kiểm hoặc lô sản xuất có thể xuất hiện nhiều dòng chi tiết lỗi. Không cộng lại số đôi kiểm/sản xuất hoặc tổng đôi lỗi trên từng dòng chi tiết. Xác định khóa bản ghi từ cấu trúc NIKE thực tế, dùng ID nguồn nếu có; không dùng một khóa suy đoán làm mất các audit khác nhau cùng ngày/model/chuyền.

HFPA lấy số lượng từng lỗi trong bản NIKE. FTT lấy số lượng lỗi trong bản NIKE, không mặc nhiên áp công thức AGrade/Rework/B-C của parser nội bộ. Nếu một đôi có nhiều lỗi, ghi rõ số lượng đang đếm lượt lỗi hay số đôi. Phần tỷ lệ DR chỉ được tính khi đã xác nhận được tử số và mẫu số đúng đơn vị từ nguồn tương ứng.

Nếu thiếu khả năng loại trùng theo đôi, không tính tỷ lệ đôi lỗi theo nhóm bằng cách dùng tổng lượt lỗi. Tổng lỗi chi tiết không bắt buộc bằng số đôi lỗi. HFPA và FTT được phân tích riêng; DR lấy từ hai nguồn này và không được cộng thêm như một quần thể thứ ba.

## 8. Công thức và cách chọn Top 3

Gọi Q(s, a, d, m) là tổng số lượng lỗi hợp lệ của nguồn s, khu a, nhóm lỗi d và model m trong kỳ. Nguồn s là HFPA NIKE hoặc FTT NIKE, luôn được tính độc lập. DR không phải nguồn s thứ ba.

| Chỉ tiêu | Công thức |
| --- | --- |
| Tổng nhóm lỗi d của khu a | Q(s,a,d) = tổng theo mọi model của Q(s,a,d,m) |
| Tổng nhóm lỗi d của CLG | Q(s,CLG,d) = Q(s,VH,d) + Q(s,JV,d) + Q(s,VB,d) |
| Đóng góp khu a vào nhóm lỗi d | Q(s,a,d) / Q(s,CLG,d) × 100% |
| Tỷ trọng nhóm lỗi d trong CLG | Q(s,CLG,d) / tổng số lượng mọi nhóm lỗi của CLG cùng nguồn × 100% |
| Tỷ trọng model trong nhóm lỗi của một khu | Q(s,a,d,m) / Q(s,a,d) × 100% |

Tỷ trọng model chỉ là thông tin bổ trợ; cột bắt buộc là đóng góp khu. Nếu còn lỗi chưa ánh xạ, tổng lỗi làm mẫu số tỷ trọng phải bao gồm số lượng chưa ánh xạ và phải nêu độ bao phủ mapping; không thu hẹp mẫu số để làm tỷ trọng đẹp hơn. Kết quả chưa đủ phân khu/ánh xạ để đảm bảo Top 3 chính xác chỉ được đánh dấu sơ bộ.

Thuật toán chính:

1. Lọc đúng nguồn, kỳ, phạm vi NIKE, bản ghi hợp lệ và danh mục khu.
2. Cộng toàn bộ model theo nhóm lỗi trước khi xếp hạng.
3. Chọn Top 3 nhóm lỗi CLG theo số lượng giảm dần. Đồng hạng dùng mã/nhãn nhóm tăng dần để kết quả ổn định, đồng thời ghi nhận có đồng hạng tại ngưỡng nếu có.
4. Giữ nguyên ba nhóm này làm ba hàng so sánh chung cho cả CLG, VH, JV, VB, giống bố cục ảnh.
5. Trong từng nhóm lỗi và từng cột khu, chọn Top 3 model riêng theo số lượng của chính nhóm lỗi đó. Đồng hạng dùng mã model hoặc tên chuẩn tăng dần. Không sao chép danh sách CLG xuống từng khu.
6. Nếu một khu có Top 3 lỗi riêng khác CLG, hiển thị note phụ từ phép xếp hạng toàn bộ lỗi của khu đó; không thay nhóm lỗi ở hàng so sánh chung.

Model ưu tiên trường có ý nghĩa tương đương model/style trong bản NIKE; cần xác minh mẫu thực tế. Giữ riêng GS, PS, TD, WMNS khi là model khác; không gộp bằng cách cắt hậu tố. Chuẩn hóa lỗi chính tả bằng danh mục có dấu vết, không fuzzy-merge tùy ý.

Mẫu số bằng 0 hiển thị N/A. Số lượng thực bằng 0 hiển thị 0. Thiếu dữ liệu hiển thị “Thiếu dữ liệu”, không biến thành 0. Xếp hạng trên số gốc, làm tròn khi hiển thị: tỷ lệ đóng góp 1 chữ số thập phân; tổng sau làm tròn có thể 99,9% hoặc 100,1%.

Không dùng đóng góp khu thay cho tỷ lệ DR. Công thức tỷ lệ DR của từng phần HFPA/FTT cần chốt trước khi triển khai; không tự dùng shipment, mẫu số khác nguồn hay số lượng từ ảnh tham chiếu. Không mặc định tử số DR là tổng lượt lỗi chi tiết.

## 9. Bố cục báo cáo

Hai nguồn HFPA NIKE và FTT NIKE có bảng phân tích riêng theo ảnh. Phần DR dùng kết quả của hai nguồn này; cách bố trí và công thức chờ xác nhận, không tạo bảng DR độc lập từ nguồn thứ ba. Nếu một slide không đủ đọc rõ, tách slide thay vì giảm font quá nhỏ.

| Nhóm lỗi và đóng góp khu | CLG | VH | JV | VB |
| --- | --- | --- | --- | --- |
| Lỗi #1; số lượng CLG; VH x%, JV y%, VB z% | Top 3 model trong lỗi #1 của CLG | Top 3 của VH | Top 3 của JV | Top 3 của VB |
| Lỗi #2; số lượng CLG; VH x%, JV y%, VB z% | Top 3 model trong lỗi #2 của CLG | Top 3 của VH | Top 3 của JV | Top 3 của VB |
| Lỗi #3; số lượng CLG; VH x%, JV y%, VB z% | Top 3 model trong lỗi #3 của CLG | Top 3 của VH | Top 3 của JV | Top 3 của VB |

Trong mỗi ô model hiển thị tên và số lượng. Màu tô theo cùng model nhất quán giữa các cột nếu dùng highlight; không cố định màu theo tên trong ảnh mẫu. Có thể dùng biểu đồ thanh chồng 100% cho đóng góp VH/JV/VB nếu giúp đọc rõ.

Tiêu đề ghi quý, từ ngày–đến ngày, loại nguồn và chế độ NIKE cho HFPA/FTT. Chân bảng ghi đơn vị, định nghĩa phần trăm và trạng thái đủ kỳ. Không mang nguyên phần IA, return-by-country hoặc giới hạn DR 0,25% trong ảnh sang nếu không có nguồn và yêu cầu tương ứng.

Nhận xét tự động chỉ nêu quan sát có số liệu: nhóm lỗi lớn nhất; khu đóng góp lớn nhất; model dẫn đầu; note khu có lỗi nổi bật khác Top 3 CLG. Không suy diễn nguyên nhân gốc, kết luận năng lực khu hoặc CAPA chỉ từ phần trăm đóng góp.

## 10. Điểm tích hợp đề xuất

Các tên dưới đây là vị trí đề xuất, chưa phải file đã triển khai.

| Thành phần | Vị trí đề xuất | Nhiệm vụ |
| --- | --- | --- |
| Điều phối | `report_tools/quarterly_defect_review/report.py` | Chạy nhánh mới và trả kết quả từng nguồn |
| Thu thập riêng | `sfc_quarterly_defect_review_acquisition.py` | HFPA NIKE, FTT NIKE, danh mục Group sandals; không đổi acquisition cũ |
| Đọc nguồn | `report_tools/quarterly_defect_review/nike_inputs.py` | Kiểm tra schema HFPA NIKE/FTT NIKE, kỳ, đơn vị, khóa và chế độ; không có DR input riêng |
| Phân khu | `report_tools/quarterly_defect_review/area_mapping.py` | Dò VH/JV/VB theo chuyền và hiệu lực |
| Phân loại lỗi | `report_tools/quarterly_defect_review/defect_mapping.py` | Quy đổi NIKE theo bảng đã xác thực |
| Tính toán | `report_tools/quarterly_defect_review/analysis.py` | Nhóm lỗi → Top model → đóng góp khu |
| Xuất | `report_tools/quarterly_defect_review/excel_output.py`, `office_output.py` | Excel nguồn, Copy/Paste sang PowerPoint |
| Đăng ký nhánh | `sfc_report_tools.py`, `report_tool_worker.py` và registry/contract/UI liên quan sau khi kiểm tra | Thêm lựa chọn mới, giữ nguyên mã và hành vi của lựa chọn cũ |

Kết quả được lưu theo run riêng, không ghi đè báo cáo hiện có. Dữ liệu phân tích nên có các trường nguồn, quý, khoảng ngày, phiên bản mapping, trạng thái đầy đủ, Top lỗi CLG, đóng góp khu và Top model theo từng lỗi/khu. Không ép dữ liệu mới vào cấu trúc `quarter.rankings` cũ.

Nếu HFPA hoặc FTT thiếu nguồn, cho phép xem phần đủ dữ liệu với nhãn rõ và đánh dấu phần DR tương ứng chưa đủ. Không yêu cầu file DR riêng; không lấy ranking nguồn khác lấp phần trống.

## 11. Excel trước, PowerPoint sau

Chuẩn bị nguồn, phép tính, bảng so sánh, biểu đồ, màu, font, nhãn và chú giải hoàn chỉnh trong Excel. Lưu và tính lại Excel trước khi chuyển nội dung sang PPTX.

Copy/Paste bảng/chart từ Excel sang PowerPoint. Với bảng trình bày có thể dùng ảnh chụp vùng Excel; chart thực hiện theo yêu cầu định dạng và khả năng Office của dự án. PowerPoint chỉ căn vị trí và kích thước. Muốn đổi nội dung hay thiết kế phải trở lại Excel và dán lại.

Lưu PPTX cuối bằng Microsoft Office hoặc WPS. Không sửa chart XML, cache, workbook nhúng hoặc liên kết trong gói PPTX sau khi lưu. Không liên kết tới workbook tạm bị xóa; giữ bản Excel cuối đi kèm khi ứng dụng tạo liên kết.

Mở lại bản xuất để kiểm tra đủ slide, không Repair, đủ bảng và nhãn không bị cắt. Ghi đúng ứng dụng thực tế đã dùng. Yêu cầu note HFPA ba tháng nếu báo cáo mới có dùng lại phải tiếp tục tạo từ đúng ba tháng được chọn trong Excel rồi dán ảnh, không cố định số theo ảnh mẫu.

## 12. Review gates và kiểm thử nghiệm thu

| Giai đoạn | Bằng chứng cần xem | Điều kiện chuyển bước |
| --- | --- | --- |
| 1. Nguồn | Mẫu HFPA NIKE, FTT NIKE và danh sách Factory Sandals của FTT | Hiểu từng cột, đơn vị và khóa; đủ dữ liệu cho kỳ được công bố; chốt công thức DR từ hai nguồn |
| 2. Mapping | Danh mục chuyền và tất cả nhãn/mã lỗi thực tế kèm số lượng | Dò duy nhất; mapping có căn cứ; phần chưa khớp được giải quyết hoặc kết quả ghi rõ chưa đầy đủ |
| 3. Tính | Bảng tổng lỗi, Top 3, Top model và đóng góp khu | Đối chiếu với phép tính độc lập/Pivot trên nguồn NIKE |
| 4. Excel | Bảng/chart hoàn chỉnh | Số và bố cục đọc rõ, tỷ lệ đúng mẫu số |
| 5. PowerPoint | Bản lưu từ Office/WPS và lần mở lại | Khớp Excel, không Repair, không liên kết tới file tạm |
| 6. Luồng cũ | Các kiểm tra hiện có cho báo cáo quý/HFPA/FTT/COPQ | Kết quả không bị thay bởi tùy chọn NIKE của báo cáo mới |

Kiểm thử nghiệp vụ quan trọng:

- CLG = VH + JV + VB cho từng nguồn, nhóm lỗi và model; không có bản ghi lọt vào hai khu.
- Với nhóm lỗi có tổng > 0 và đầy đủ phân khu, đóng góp VH + JV + VB bằng 100% trước làm tròn.
- Một chuyền sandals gốc VH được chuyển sang VB thì tổng CLG giữ nguyên; VH giảm đúng phần VB tăng.
- Hai nhà máy có cùng Line 1 không bị dò nhầm; chuyền đổi Group giữa quý áp đúng hiệu lực.
- Cả ba nhóm lỗi được chọn từ toàn CLG, không lấy từ lỗi của model đứng đầu; ba cột khu được xếp hạng model độc lập.
- HFPA và FTT có fixture khác nhau để chứng minh không dùng chung Counter/ranking hay fallback chéo nguồn; phần DR truy được về đúng nguồn tương ứng.
- HFPA không có cột Group vẫn phân khu VB đúng nhờ danh mục Factory Sandals của FTT; phép dò không thay số lượng lỗi HFPA bằng số lượng FTT.
- Có trường hợp nhiều lỗi trên một audit/lô để kiểm tra không cộng trùng mẫu số và số đôi lỗi.
- Đủ 1, 2 hoặc ít hơn 3 model/lỗi thì hiển thị đúng số thực có; không điền model giả.
- Nhãn lỗi chưa ánh xạ, model thiếu hoặc chuyền chưa dò được giữa danh mục FTT và HFPA phải xuất hiện trong bảng review.
- File nội bộ cùng tên/kỳ không được cache thay file NIKE; chạy báo cáo mới rồi chạy nhánh cũ không làm nhánh cũ bật NIKE.
- Kỳ thiếu tháng 8 không được ghi “Q1 đầy đủ”; không cộng bình quân các phần trăm tháng thay cho tỷ lệ từ tổng số lượng.

## 13. Thứ tự triển khai

1. Xác nhận hai mẫu xuất NIKE, công thức DR và danh sách chuyền Group Factory Sandals của FTT. Hoàn thiện hợp đồng dữ liệu và mapping thực tế trước.
2. Tạo module riêng, chuẩn hóa dữ liệu và audit phân khu/phân loại lỗi; giữ nguyên luồng cũ.
3. Tính kết quả và đối chiếu thủ công một nhóm lỗi, một model, một chuyền sandals ở cả HFPA/FTT; đối chiếu phần DR với đúng nguồn NIKE tương ứng.
4. Tạo bảng và chart trong Excel, sau đó nối xuất PowerPoint qua Office/WPS.
5. Thêm lựa chọn vào tool quý và kiểm tra cách ly nguồn/cache, kiểm tra hồi quy các luồng đang chạy.
6. Chạy đủ FY27 Q1 khi có nguồn hợp lệ; gửi người dùng bảng Top 3, đóng góp khu và những dòng cần review.

Thông tin còn cần để tính số thật: dữ liệu HFPA NIKE và FTT NIKE đủ Q1, danh sách chuyền đã được group Factory Sandals ở FTT, cùng xác nhận công thức DR. Không cần file DR riêng. Nếu công cụ có phiên SFC đang đăng nhập, có thể thực hiện bước thu thập bằng nhánh riêng sau khi triển khai; chưa coi các file nội bộ tìm thấy là đầu vào hợp lệ.
