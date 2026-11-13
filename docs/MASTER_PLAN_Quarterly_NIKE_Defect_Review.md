# Master plan cho Anti Gravity IDE

Phân tích lỗi quý từ HFPA NIKE và FTT NIKE. Bản chốt để triển khai ngày 11/11/2026.

## 1. Kết quả cần xây dựng

Tích hợp một nhánh báo cáo mới vào Chingluh Dashboard tại workspace `C:\Users\User\Documents\Chingluh Dashboard`. Nhánh này phân tích HFPA và FTT riêng biệt từ hai bản xuất NIKE, chia khu CLG/VH/JV/VB, chọn Top 3 nhóm lỗi và phân tích sâu những lỗi thực tế nằm trong từng nhóm.

Điểm bổ sung quan trọng đã được người dùng xác nhận: báo cáo hiện tại chỉ trình bày nhóm lỗi tổng quát ở cột HFPA như Bonding. Báo cáo mới phải cho biết bên trong Bonding là những lỗi cụ thể nào, và Top 3 lỗi chi tiết có số lượng lớn nhất là gì. Không dừng ở việc giải thích định nghĩa Bonding hoặc liệt kê các nhóm con lý thuyết.

Giữ cả hai kết quả song song: Top 3 lỗi chi tiết trong mỗi nhóm và Top 3 model trong mỗi nhóm. Model là chiều phân tích riêng; không thay phần Top model bằng phần Top lỗi chi tiết.

Deliverables khi IDE hoàn thành: mã nguồn nhánh mới, mapping có dấu vết, kiểm thử nghiệp vụ, workbook Excel nguồn có bảng/chart hoàn chỉnh, PPTX được lưu qua Office/WPS, báo cáo đối chiếu và hướng dẫn sử dụng. Master plan này là tài liệu bàn giao, không phải bằng chứng tính năng đã được triển khai hoặc số Q1 đã được xác minh.

## 2. Phạm vi và giới hạn bắt buộc

| Nội dung | Quy tắc triển khai |
| --- | --- |
| Nguồn số liệu | Chỉ HFPA NIKE và FTT NIKE cho nhánh báo cáo mới |
| Luồng đang chạy | Giữ nguyên HFPA tháng, FTT/COPQ, QAStation, MES410, báo cáo quý hiện tại và cấu hình mặc định của chúng |
| Tableau dùng chung | Không append/ghi đè bằng dữ liệu của báo cáo mới; không dùng cache nội bộ để lấp nguồn NIKE |
| DR | Người dùng xác nhận lấy từ HFPA NIKE và FTT NIKE; không yêu cầu nguồn/file DR riêng, không mặc định là hàng trả về |
| Công thức DR | Chưa xác nhận đầy đủ tên gọi, tử số và mẫu số. Tách cấu hình để chốt sau; triển khai Top lỗi/Top model/phân khu trước, không tự đặt công thức |
| Kỳ | Mặc định dự kiến FY27 Q1 = 01/06/2026-31/08/2026 theo template; hiển thị rõ ngày để người dùng sửa nếu cần |
| Ảnh báo cáo Q2 | Chỉ tham chiếu bố cục; không dùng số lượng, tỷ lệ, giới hạn, tháng hoặc tên model trong ảnh để làm dữ liệu Q1 |
| Chỉnh sửa mã đang có | Kiểm tra git diff trước khi làm; không reset, ghi đè hay dọn thay đổi của người dùng |

Không cần phát triển một hệ thống báo cáo khác hoặc viết lại các engine cũ. Thêm module có phạm vi riêng, nối vào bộ chạy tool hiện có bằng thay đổi tối thiểu. Không sửa mặc định NIKE toàn cục.

## 3. Phân khu và nguồn danh mục VB

| Khu | Phạm vi đã chốt |
| --- | --- |
| VH | VH + VH2, loại tất cả chuyền thuộc Group Factory Sandals |
| JV | JV + JV2 |
| VB | Toàn bộ chuyền đã thuộc Group Factory Sandals ở FTT |
| CLG | VH + JV + VB; một bản ghi chỉ thuộc một khu |

FTT có Group Factory Sandals; HFPA không hiển thị Group. Lấy danh sách chuyền đã group sẵn trong FTT làm danh mục để dò sang HFPA. Việc HFPA không có cột Group không phải lỗi thiếu schema và không được dùng làm lý do bỏ phần VB.

FTT cung cấp danh mục phân khu. Số lượng HFPA phải lấy từ HFPA NIKE; số lượng FTT phải lấy từ FTT NIKE. Tuyệt đối không sao chép lượng lỗi FTT sang HFPA khi dò chuyền.

Danh mục cần nhà máy, xưởng, chuyền gốc, mã chuẩn hóa, Group gốc, hiệu lực và nguồn chứng minh. Dò theo nhà máy + xưởng + chuyền, thêm khoảng hiệu lực nếu chuyền đổi Group trong kỳ. Không dò bằng số chuyền đơn lẻ hoặc tên model có chữ Slide/Sandal.

Chuẩn hóa có kiểm soát: chữ hoa/thường, khoảng trắng và cách viết Line1/Line 1. Giữ mã gốc, không bỏ hậu tố hoặc ký tự phân biệt chuyền. Nếu NIKE chỉ có mã ghép, tách theo cấu trúc đã xác minh từ mẫu thực tế.

Trình tự: đọc danh mục FTT, đánh dấu chuyền sandals, phân khu bản ghi từng nguồn, loại sandals khỏi VH, kiểm tra đối chiếu tổng. Nếu Group sandals chứa chuyền JV/JV2 gây chồng quy tắc đã chốt, báo đúng trường hợp xung đột để xác nhận; không tự tính hai lần. Bản ghi thiếu khóa hoặc khớp nhiều khu phải được đưa vào bảng review.

Không cộng bản xuất FTT toàn bộ với bản xuất Factory Sandals như hai quần thể độc lập. Bản Group là danh mục/đối chiếu; nếu dùng làm nguồn lượng phải loại phần đã nằm trong tổng.

## 4. Mô hình phân loại lỗi: phải giữ đủ độ chi tiết

| Tầng dữ liệu | Ý nghĩa | Ví dụ theo ảnh |
| --- | --- | --- |
| Nhóm chuẩn HFPA | Nhóm tổng quát dùng để chọn Top 3 nhóm lỗi | Bond gap / Bonding |
| Loại theo cột FTT NIKE | Phân loại trung gian trong bảng quy đổi | Bond gap / Rat hole; Delamination |
| Lỗi chi tiết thực tế | Tên/mã lỗi quan sát được trong bản NIKE, có lượng tương ứng | Chỉ lấy nhãn có thật trong nguồn, không tự tạo từ ví dụ mô tả |
| Model và khu | Các chiều phân tích lượng lỗi | Model thực tế; CLG/VH/JV/VB |

Trong ảnh mapping, Bonding chỉ có hai loại ở cột FTT. Yêu cầu Top 3 mới là Top 3 tên lỗi chi tiết thực tế bên trong Bonding nếu nguồn có đủ chi tiết, không phải ép hai loại trung gian thành ba loại.

Cột Description trong ảnh giải thích phạm vi phân loại. Những ví dụ như outsole to midsole hoặc outsole to upper chỉ là mô tả; không được coi là các lỗi đã phát sinh, không được tạo lượng, thứ hạng hay phân bổ từ các ví dụ này khi dữ liệu chưa ghi nhận chúng.

Không gộp mất nhãn chi tiết sau khi đã quy về nhóm chuẩn. Luôn giữ ít nhất: nhóm chuẩn, loại FTT nếu xác định được, nhãn/mã chi tiết gốc và nhãn chi tiết chuẩn hóa.

Chỉ tách một nhãn thành nhiều lỗi khi nguồn có số lượng riêng cho từng lỗi. Không nhân toàn bộ lượng của một dòng cho từng cụm từ hoặc từng vị trí mô tả. Nếu một dòng liệt kê nhiều vị trí mà không phân lượng, giữ nguyên tổ hợp gốc hoặc báo chưa thể tách.

## 5. Bảng quy đổi nhóm chuẩn theo ảnh NIKE

| Nhóm chuẩn HFPA | Các loại FTT tương ứng |
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

Đối chiếu lại cách viết với mẫu NIKE thực tế. X-ray được đặt theo ô nhóm Wrinkle / Crease / Mis-shape trong ảnh. Có thể hiển thị nhãn rút gọn Bonding, nhưng ID nhóm chuẩn phải ổn định.

Mapping gồm source_metric, original_code/label, standard_group_id, nike_subtype_id nếu có, detail_defect_id/label, phiên bản, căn cứ và người/thời điểm xác nhận khi có điều chỉnh. Ưu tiên mã lỗi rõ ràng, sau đó exact label/alias đã xác minh. Không dùng fuzzy matching hoặc AI tự gán làm kết quả chính thức.

Lỗi chưa có mapping phải hiển thị “Chưa ánh xạ” và số lượng ảnh hưởng. Không gán mặc định vào Other/Miscellaneous. Nhóm Other chính thức vẫn được xếp hạng như các nhóm khác. Một lỗi chi tiết có thể cùng tên ở hai nhóm; khóa cần bao gồm nhóm cha và nguồn, không chỉ tên.

## 6. Khảo sát hai bản NIKE trước khi viết adapter

IDE phải lấy/kiểm tra mẫu HFPA NIKE và FTT NIKE đúng nhà máy, kỳ, bộ lọc và chế độ. Xác nhận schema thật; không hard-code schema nội bộ làm schema NIKE.

Lập báo cáo schema gồm tên sheet/header row, tên cột gốc, trường chuẩn tương ứng, kiểu, đơn vị, khóa, dòng tổng, độ hạt và mức chi tiết lỗi có thể lấy. Ghi nhận chế độ NIKE trước/sau xuất; tên file có chữ NIKE chưa đủ chứng minh nguồn.

Ba mức hỗ trợ chi tiết phải phân biệt:

- Có tên lỗi thực tế và lượng theo từng lỗi: tính Top lỗi chi tiết.
- Chỉ có loại trung gian FTT và lượng: hiển thị theo mức nguồn đang có, nêu rõ chưa đủ lỗi chi tiết; không gọi đó là kết quả đầy đủ theo yêu cầu mới.
- Chỉ có nhóm tổng HFPA: tính được nhóm lớn/đóng góp khu/Top model nếu đủ trường, nhưng phần Top lỗi chi tiết phải ghi chưa có dữ liệu chi tiết.

Nếu HFPA NIKE không có chi tiết mà FTT NIKE có, chỉ FTT có thể có Top lỗi chi tiết đầy đủ. Dùng bảng mapping từ FTT để phân loại tên lỗi có thật trong HFPA khi phù hợp; không dùng lượng FTT để chia nhỏ tổng HFPA. Không âm thầm xuất nội bộ hoặc thêm một nguồn ngoài phạm vi để bù.

Có thể tiếp tục các phần độc lập khi một mục thiếu dữ liệu, nhưng báo cáo phải ghi rõ hạn chế. Chỉ hỏi người dùng về trường/mapping hoặc nghiệp vụ thực sự chưa xác định, kèm mẫu dữ liệu nhỏ có liên quan. Không hỏi lại các quy tắc đã chốt trong tài liệu này.

## 7. Hợp đồng dữ liệu và kiểm soát cộng trùng

Tách bảng base_records và defect_facts để không cộng đôi kiểm/sản xuất trên mỗi dòng chi tiết lỗi.

| Bảng | Các trường tối thiểu |
| --- | --- |
| Source manifest | run_id, metric, report_mode=NIKE, factory, date_from/date_to, filter_snapshot, file hash, schema version, coverage |
| Line catalog | factory, plant, line_raw, line_std, source_group, report_area, effective_from/to, provenance |
| Base records | source_record_id, metric, date, factory/plant/line, model key/label, inspected/produced quantity, defective-pair quantity nếu có |
| Defect facts | base_record_id, source_file/sheet/row/column, defect_group_id, nike_subtype_id, detail_raw/id/label, issue_quantity, quantity_unit, report_area, detail_status |
| Mapping review | raw label/key, suggested or confirmed mapping, affected quantity, status, reason, provenance |

Xác minh grain trước khi dedup. Dùng ID nguồn nếu có; nếu không có phải chứng minh khóa tổng hợp không làm mất hai audit/lô khác nhau cùng ngày/chuyền/model. Giữ bằng chứng dòng nguồn và ghi rõ dòng bị loại vì trùng export hay dòng tổng.

Không đồng thời cộng subtotal nhóm cha, subtotal loại FTT và lượng lỗi chi tiết như ba lần lỗi. Chọn lớp chi tiết nguyên tử có ý nghĩa từ nguồn để cộng; lớp tổng dùng làm control. Nếu lượng nhóm cha lớn hơn tổng lỗi chi tiết có thể đối chiếu, đưa chênh lệch vào “Chưa có chi tiết” để thể hiện độ bao phủ, không coi phần chênh là một lỗi thực tế để tranh Top 3.

Nếu lượng chi tiết vượt control khi hai lớp được nguồn khẳng định cùng đơn vị và quần thể, báo không khớp và chặn kết quả hoàn tất. Nếu một lớp là đôi lỗi và một lớp là lượt lỗi thì không ép bằng nhau; nêu khác đơn vị và không suy ra tỷ lệ đôi lỗi từ lượt lỗi.

Mỗi đầu vào cache phải phân biệt loại báo cáo, metric, NIKE, nhà máy, khoảng ngày, Group/trạm, schema, mapping và catalog version. Cùng nội dung xuất lặp không được cộng hai lần. Thay đổi mapping phải tính lại kết quả phụ thuộc.

## 8. Quy tắc xếp hạng và công thức

Tính độc lập theo metric HFPA/FTT và scope CLG/VH/JV/VB. Gọi G(s,a,g) là lượng nhóm lỗi g, D(s,a,g,d) là lượng lỗi chi tiết d, M(s,a,g,m) là lượng model m trong nhóm g. Tất cả cùng kỳ, cùng đơn vị phù hợp và không cộng trùng.

| Chỉ tiêu | Công thức hoặc quy tắc |
| --- | --- |
| Lượng CLG của nhóm g | G(s,VH,g) + G(s,JV,g) + G(s,VB,g) |
| Đóng góp khu vào nhóm g | G(s,a,g) / G(s,CLG,g) x 100% |
| Top 3 nhóm lỗi | Xếp G(s,CLG,g) giảm dần, lấy 3 nhóm; tính từ mọi model |
| Top 3 lỗi chi tiết trong nhóm | Trong từng nhóm g và scope a, xếp D(s,a,g,d) giảm dần, lấy tối đa 3 lỗi thực tế |
| Tỷ trọng lỗi chi tiết trong nhóm | D(s,a,g,d) / G(s,a,g) x 100%, chỉ khi tử/mẫu cùng đơn vị/quần thể và lượng nhóm có thể đối chiếu |
| Đóng góp khu vào cùng lỗi chi tiết | D(s,a,g,d) / D(s,CLG,g,d) x 100% |
| Top 3 model của nhóm | Xếp M(s,a,g,m) giảm dần trong chính nhóm g và scope a |
| Độ bao phủ chi tiết | Tổng lượng chi tiết đối chiếu được / lượng nhóm cùng đơn vị; không dùng như chỉ số chất lượng sản phẩm |

Phải phân biệt ba tỷ lệ: đóng góp khu vào nhóm, tỷ trọng lỗi con trong nhóm tại một khu và đóng góp khu vào một lỗi con. Ghi rõ tên mẫu số trên bảng/tooltip/chân trang; không gọi cả ba là một cột “%”.

Ba nhóm lỗi CLG là ba hàng chung cho mọi khu trong bảng overview. Top lỗi chi tiết và Top model của từng khu được xếp riêng trong cùng nhóm, không sao chép CLG xuống các khu. Có thể thêm note về lỗi đứng đầu riêng của khu nếu khác Top nhóm CLG.

Top 3 lỗi chi tiết không dùng mẫu số tổng Top 3; mẫu số là toàn nhóm tương ứng, gồm các lỗi còn lại và phần chưa có chi tiết có thể đối chiếu. Với coverage chưa đủ, ghi “Top trong phần đã có chi tiết”; không khẳng định Top 3 thật của toàn nhóm. Unknown group/area làm ảnh hưởng kết quả cũng phải được công bố hoặc giải quyết trước khi báo hoàn tất.

Nếu chỉ có 1 hoặc 2 lỗi thực tế thì hiển thị 1 hoặc 2. Không tạo lỗi thứ ba, không tách hai loại Bonding thành ba bằng suy diễn. Đồng hạng sắp theo ID/nhãn chuẩn tăng dần và ghi nhận đồng hạng ở ngưỡng. Số 0 khác thiếu dữ liệu; mẫu số 0 hiển thị N/A. Tính trên số gốc, phần trăm hiển thị 1 số thập phân.

Model dùng trường đã xác minh của bản NIKE. Theo hướng dẫn bàn giao hiện có, ưu tiên Model cho FTT nếu bản NIKE có trường đó; không đổi sang Model 2 tùy ý. Nếu schema khác, ghi mapping tương đương và chốt từ mẫu. Không gộp GS/PS/TD/WMNS chỉ bằng xóa hậu tố.

DR chưa chốt công thức thì phần tỷ lệ DR phải có trạng thái chưa xác nhận. Không lấy shipment từ ảnh, không cộng HFPA + FTT, không gọi tỷ trọng đóng góp khu là DR. Việc này không chặn triển khai các bảng số lượng và Top lỗi đã đủ dữ liệu.

## 9. Ví dụ kiểm thử tính toán có đáp án

Dữ liệu sau chỉ là fixture giả lập để kiểm tra engine, không phải số Q1. Tên A/B/C/D đại diện các nhãn lỗi thực tế mà fixture cung cấp; không phải taxonomy Bonding mới.

| Chi tiết thuộc Bonding | VH | JV | VB | CLG |
| --- | --- | --- | --- | --- |
| Lỗi A | 300 | 200 | 150 | 650 |
| Lỗi B | 70 | 110 | 70 | 250 |
| Lỗi C | 20 | 30 | 20 | 70 |
| Lỗi D | 10 | 10 | 10 | 30 |
| Tổng Bonding | 400 | 350 | 250 | 1.000 |

Đáp án: đóng góp VH/JV/VB vào Bonding lần lượt 40,0% / 35,0% / 25,0%. Top chi tiết CLG là A=650 (65,0%), B=250 (25,0%), C=70 (7,0%); phần còn lại D=3,0%. Tỷ trọng A trong Bonding của VH là 300/400=75,0%. Đóng góp VH vào lỗi A của CLG là 300/650=46,2%. Hai tỷ lệ 75,0% và 46,2% không được tráo nhau.

Tạo fixture riêng cho HFPA và FTT có thứ hạng khác nhau để bắt lỗi dùng chung kết quả. Thêm ví dụ chỉ có hai lỗi Bonding, tổng nhóm có 200 lỗi chưa có chi tiết, một lỗi có nhiều vị trí không chia lượng và một chuyền sandals có cùng số chuyền với nhà máy khác.

## 10. Thiết kế báo cáo và tương tác

Mỗi metric có hai vùng trình bày chính, có thể tách slide để đủ đọc rõ.

| Vùng | Nội dung |
| --- | --- |
| Overview theo ảnh | Top 3 nhóm lỗi CLG, lượng và đóng góp VH/JV/VB; Top 3 model trong mỗi nhóm cho CLG/VH/JV/VB |
| Phân tích chi tiết mới | Trong từng nhóm trên: Top 3 lỗi thực tế, lượng, tỷ trọng trong nhóm, coverage; CLG và từng khu có kết quả riêng |
| Drill-down | Nhấp/chọn nhóm và khu để xem mọi lỗi chi tiết, loại FTT, model, chuyền và dòng nguồn liên quan |
| Review dữ liệu | Kỳ/phạm vi thiếu, mapping chưa khớp, chuyền chưa dò được, chênh lệch control và công thức DR chưa chốt |

Ví dụ bố cục phần chi tiết: hàng là Bonding, Cement hoặc nhóm Top 3 thật của kỳ; cột là CLG/VH/JV/VB; trong từng ô liệt kê tối đa 3 tên lỗi thực tế, lượng và % trong nhóm của scope đó. Nhãn phải đầy đủ, xuống dòng hợp lý. Đóng góp khu vào lỗi con có thể xem ở bảng drill-down riêng để tránh trộn mẫu số.

Không nhân bảng thành 3 nhóm x 3 lỗi x 3 model trên một slide. Top model của nhóm vẫn giữ bảng tổng quan; nếu cần truy vấn model của riêng lỗi con thì thực hiện trong drill-down, không thay đổi định nghĩa Top model ban đầu.

Giao diện cần chọn quý, ngày chốt, nguồn/mode riêng, trạng thái NIKE, danh mục sandals, kiểm tra đầu vào, xem preview kết quả và chạy xuất. Hiển thị đúng trạng thái “Không có lỗi”, “Thiếu dữ liệu chi tiết” hoặc “Chưa ánh xạ”. Một lỗi vắng ở khu đầy đủ dữ liệu khác với một khu thiếu dữ liệu.

Không dùng số/model/màu highlight cố định từ ảnh Q2. Nhận xét tự động chỉ nêu số liệu: lỗi con đứng đầu, khu đóng góp lớn nhất, model liên quan và coverage. Không suy diễn root cause/CAPA từ tên lỗi hoặc tỷ trọng đóng góp.

## 11. Kiến trúc tích hợp đề xuất

Đặt mã trong module riêng `report_tools/quarterly_defect_review/`. Các tên dưới đây là đề xuất; IDE kiểm tra kiến trúc hiện tại để nối đúng điểm mà không viết lại luồng cũ.

| File/thành phần | Trách nhiệm |
| --- | --- |
| report.py | Điều phối validate, normalize, map, analyze, export; trả status theo phạm vi |
| nike_inputs.py | Adapter riêng HFPA NIKE/FTT NIKE, schema profiles và provenance |
| area_mapping.py | Catalog Factory Sandals, dò chuyền, phân VH/JV/VB |
| defect_mapping.py + config | Bảng nhóm chuẩn/loại FTT/lỗi chi tiết và danh sách chưa khớp |
| analysis.py | Top nhóm, Top chi tiết, Top model, đóng góp khu và coverage |
| contracts.py | Dữ liệu typed và validation; không tái định nghĩa các khóa cũ toàn cục |
| excel_output.py | Workbook nguồn, công thức/bảng/chart và thiết kế hoàn chỉnh |
| office_output.py | Copy/Paste đối tượng Excel, căn vị trí/kích thước, lưu và mở lại PPTX |
| sfc_quarterly_defect_review_acquisition.py | Thu thập riêng NIKE và danh mục Group; tái dùng helper an toàn bằng tham số riêng |
| tests/test_quarterly_defect_review_*.py | Fixture và kiểm thử nghiệp vụ của nhánh mới |

Điểm nối đã xác định trong dự án: `report_registry.json`, `report_contracts.json`, `sfc_report_tools.py`, `report_tool_worker.py`, `templates/report_tools.html`, `static/report_tools.js` và các phần validation/export thực tế cần thiết. Thêm mã báo cáo `quarterly_defect_review`; không đổi runner/nguồn của `quarterly_quality` đang có. Việc đăng ký UI/registry phải tuân theo schema của phiên bản mã hiện tại.

Các năng lực hiện có đáng kiểm tra để tái sử dụng có kiểm soát: `sfc_hfpa.py` và `sfc_excel.py` hỗ trợ xuất/validate HFPA NIKE; module quarterly_quality có logic quý và adapter Office. Không gọi trực tiếp parser nội bộ với file NIKE chỉ vì cột có vẻ giống nhau. Không tái sử dụng `quarter.rankings` cũ nếu mất chiều nhóm/lỗi chi tiết/khu.

API tính toán nên tách rõ normalize_nike_source, build_area_catalog, classify_areas, map_defects, aggregate_defects, rank_groups, rank_details, rank_models và build_review. Các hàm tính thuần không cần SFC hay Office để unit test.

## 12. Xuất Excel và PowerPoint đúng quy tắc dự án

Đọc AGENTS.md ở workspace và mọi chỉ dẫn áp dụng trước khi code. Nguồn, dữ liệu, công thức, series, trục, màu, font, nhãn, legend và toàn bộ thiết kế chart phải hoàn thiện trong Excel trước.

Workbook nên có phần overview, chi tiết nhóm, dòng lỗi chuẩn hóa, mapping/catalog và review cần thiết; có thể gộp các vùng khi đơn giản. Nguồn gốc được giữ nguyên hoặc tham chiếu file nguồn kèm manifest. Không tạo quá nhiều tab chỉ lặp số đã tính. Mọi kết quả trên slide truy về vùng Excel tương ứng.

Copy/Paste chart từ Excel bằng Office/WPS sang PowerPoint. Giữ yêu cầu chart là đối tượng native từ bản bàn giao hiện có; không thay chart bằng PNG/EMF hoặc tự dựng chart trong PPTX. Bảng trình bày có thể dán ảnh vùng Excel khi phù hợp. Sau dán PowerPoint chỉ chỉnh vị trí/kích thước, không sửa nội dung/dữ liệu/format chart.

Lưu Excel nguồn cuối ở đường dẫn giao kèm trước khi Copy/Paste. Dùng Office/WPS lưu PPTX cuối, không vá chart XML/cache/workbook nhúng/relationships sau lưu. Không để liên kết tới Excel tạm đã bị xóa. Nếu adapter hiện tại có lỗi WPS hoặc thiếu chức năng, báo đúng lỗi và tiếp tục sửa trong phạm vi nhánh mới; không bỏ qua bằng XML.

Mở lại bản xuất bằng ứng dụng thực tế, kiểm tra không Repair, đủ slide, nhãn đầy đủ và đối chiếu Excel. Ghi rõ đã kiểm tra bằng Microsoft Office hay WPS. Nếu tái dùng note HFPA bên phải slide 1, tạo vùng bảng từ đúng ba tháng của quý trong Excel, tính lại rồi chụp vùng dán sang PPTX.

## 13. Trạng thái chạy và audit trail

Mỗi run lưu input manifest, hash nguồn, export mode/bộ lọc, phiên bản schema/mapping/catalog, độ bao phủ theo nguồn/nhà máy/tháng, log kiểm tra và output paths. Không sửa dữ liệu nguồn khi người dùng review alias; chỉnh mapping tách biệt và tính lại.

Các trạng thái chính: đang kiểm tra, thiếu nguồn, có kết quả một phần, cần review mapping, sẵn sàng xuất, xuất hoàn tất hoặc lỗi xuất. Không dùng một cờ passed duy nhất để che trường thiếu chi tiết.

Chưa đủ quý phải ghi ngày chốt thực tế, không gọi đủ Q1. File rỗng không chứng minh không có lỗi nếu chưa xác thực export/filter và coverage. Khi thiếu chi tiết, cho phép báo cáo nhóm lớn nhưng ghi phần nào chưa đáp ứng Top lỗi con.

Phần DR chưa xác nhận công thức phải hiện rõ trong preview/review; không tự đặt số 0 hoặc dùng phần trăm từ ảnh. Không cần chờ DR để làm xong nhánh nhóm/chi tiết/model khi các phần đó có nguồn hợp lệ.

## 14. Kiểm thử nghiệm thu bắt buộc

| Trường hợp | Kết quả cần chứng minh |
| --- | --- |
| Fixture Bonding 1.000 | Đúng các kết quả 40/35/25, 65/25/7, 75,0 và 46,2% ở mục 9 |
| Nhóm có 2 lỗi thực tế | Chỉ hiển thị 2, không tạo thứ ba từ Description |
| Nguồn chỉ có HFPA group | Báo thiếu chi tiết; không chia bằng số FTT |
| HFPA không có Group | VB được dò từ danh mục FTT; lượng HFPA giữ nguyên nguồn |
| Chuyền sandals gốc VH | VH giảm đúng phần VB tăng; CLG không đổi, không tính trùng |
| Cùng Line 1 khác nhà máy | Không khớp nhầm; kiểm tra cả xưởng và hiệu lực |
| Parent/subtype/detail cùng xuất hiện | Không cộng nhiều lớp hoặc subtotal vào tổng lỗi |
| Nhiều vị trí trong một dòng | Không nhân số lượng đầy đủ cho từng vị trí |
| Audit/lô lặp qua nhiều lỗi | Mẫu số kiểm/sản xuất và đôi lỗi không bị cộng nhiều lần |
| Coverage chi tiết chưa đủ | Mẫu số không thu hẹp còn Top 3; công bố coverage và lượng chưa có chi tiết |
| Unknown mapping | Hiện review; không gán ngầm vào Other |
| Xếp hạng theo khu | Có fixture thứ hạng khác CLG để phát hiện sao chép Top 3 |
| HFPA khác FTT | Kết quả tính độc lập; không chia sẻ Counter/cache sai nguồn |
| Đồng hạng, thiếu, 0 | Thứ tự ổn định; N/A/thiếu/0 đúng ý nghĩa |
| Mode/cache | File nội bộ không thay file NIKE; thay mapping/catalog làm mới kết quả phụ thuộc |
| Cách ly luồng cũ | Chạy nhánh mới không đổi mặc định hay nội dung nguồn dùng chung; các kiểm tra cũ liên quan vẫn qua |
| Bản xuất Office/WPS | Excel và PPT khớp; mở lại không Repair, không có liên kết tạm |

Đối chiếu tính toán với Pivot hoặc phép tính độc lập trên cùng nguồn NIKE, không chỉ so workbook với chính JSON đã dùng để ghi workbook. Kiểm tra đủ ba cấp: dòng nguồn, tổng nhóm và ô kết quả Top chi tiết. Chỉ chạy các test phù hợp thay đổi; không lặp vô hạn các kiểm tra đã qua.

## 15. Các giai đoạn IDE cần thực hiện

1. Khảo sát: đọc chỉ dẫn, git status/diff, registry/runner/UI, schema NIKE thật và danh sách Group sandals. Ghi các khả năng chi tiết và thiếu dữ liệu có thật.
2. Chuẩn hóa: xây hợp đồng dữ liệu, tách base/defect facts, danh mục chuyền, mapping 3 tầng có raw labels và review không khớp.
3. Engine: tính Top nhóm, Top lỗi chi tiết, Top model, phần trăm và coverage. Chạy fixture có đáp án trước khi kết nối Office.
4. Dữ liệu thật: đối chiếu một chuyền sandals giữa hai nguồn và một nhóm lỗi về tận dòng gốc; xác minh lượng không trùng và chi tiết không bị suy diễn.
5. UI: thêm nhánh riêng, chọn kỳ/nguồn, preview overview và chi tiết, bảng review, tải kết quả.
6. Excel/PowerPoint: thiết kế trong Excel, Copy/Paste native, lưu và mở lại qua Office/WPS. Tách slide khi bảng dài.
7. Hồi quy và bàn giao: kiểm tra luồng cũ, ghi phạm vi thay đổi, test đã chạy, ứng dụng đã dùng, hạn chế còn thật và cách xuất lại bằng tool.

Không dừng sau khi dựng mock UI hoặc một file báo cáo viết tay. Mục tiêu IDE là tính năng chạy lại được bằng dữ liệu và quý khác mà không sửa code hoặc hard-code model/lỗi.

## 16. Điều kiện hoàn tất và phần còn cần xác nhận

Tính năng hoàn tất khi người dùng có thể chạy nhánh mới, phân khu đúng, xem Top 3 nhóm cùng Top 3 lỗi thực tế bên trong từng nhóm và Top 3 model, kiểm tra mẫu số rõ ràng, xuất Excel/PPT qua đúng quy trình và truy nguồn được kết quả. Những phần thiếu dữ liệu phải hiển thị trạng thái rõ, không được báo đã đáp ứng đầy đủ.

Các quy tắc đã chốt không cần hỏi lại: NIKE chỉ cho báo cáo này; FTT cung cấp Group sandals để dò HFPA; VH loại sandals; JV gồm JV/JV2; CLG gồm VH/JV/VB; lỗi trước rồi model; Top chi tiết là lỗi thực tế; phần trăm khu là đóng góp vào cùng lỗi CLG; không tạo lỗi thứ ba khi chỉ có hai.

Còn cần xác minh từ mẫu thật: schema/độ chi tiết từng bản NIKE, khóa truy audit/lô, đơn vị lượng lỗi, độ đầy đủ Q1, trường model tương đương và danh mục chuyền đúng kỳ. Còn cần người dùng chốt công thức DR trước khi triển khai chỉ tiêu tỷ lệ đó. Hỏi cụ thể khi đã chỉ ra trường còn thiếu; tiếp tục các phần độc lập có đủ thông tin.

## 17. Lệnh bàn giao có thể dán vào Anti Gravity IDE

Hãy triển khai master plan này trong workspace Chingluh Dashboard. Đọc AGENTS.md, các chỉ dẫn áp dụng và kiểm tra thay đổi hiện có trước khi sửa. Tạo nhánh báo cáo riêng quarterly_defect_review, không đổi nguồn, cache hay mặc định của các luồng hiện tại.

Dùng HFPA NIKE và FTT NIKE cho riêng báo cáo này. Lấy danh sách chuyền Group Factory Sandals của FTT để phân VB cho cả hai nguồn dù HFPA không có Group. Mỗi nguồn giữ số liệu của chính nó. Phân VH/JV/VB và kiểm tra CLG không cộng trùng.

Điểm bắt buộc mới: giữ nhóm lỗi HFPA làm nhóm cha, bảo toàn loại FTT và lỗi chi tiết thực tế. Với Top 3 nhóm CLG, tính riêng Top 3 lỗi chi tiết và Top 3 model ở CLG/VH/JV/VB. Không chỉ liệt kê hai loại Bonding trong mapping; phải xếp hạng lỗi thực tế có lượng trong nguồn. Không suy diễn chi tiết từ Description hoặc dùng lượng FTT phân bổ cho HFPA. Thiếu chi tiết thì báo đúng phần thiếu, không tự xuất nội bộ.

Thực hiện lần lượt khảo sát schema, hợp đồng dữ liệu, mapping, engine, kiểm thử fixture, đối chiếu nguồn, UI và xuất Office/WPS. Hoàn thiện dữ liệu cùng thiết kế trong Excel rồi Copy/Paste sang PowerPoint. Không sửa chart hoặc XML/cache PPTX sau khi lưu. Chạy test phù hợp, mở lại bản xuất và ghi đúng ứng dụng đã kiểm tra.

Không yêu cầu file DR riêng. Phần DR lấy từ hai nguồn NIKE; khi công thức chưa được xác nhận, giữ trạng thái chưa chốt cho tỷ lệ DR và tiếp tục các chức năng độc lập. Khi bàn giao, báo file đã đổi, chức năng đã chạy được, bằng chứng test/đối chiếu, đường dẫn output và các câu hỏi còn cần xác nhận. Không chỉ trả về một kế hoạch mới.
