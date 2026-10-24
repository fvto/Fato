# DEVEXPRESS AUTOMATION & AUDIT DATA RECONCILIATION RULES

## 1. DevExpress ASP.NET WebForms Automation Guardrails
- **Accuracy over Speed**: Độ chính xác và toàn vẹn của dữ liệu là ưu tiên số một. Chấp nhận tăng thời gian chờ để loại bỏ 100% rủi ro dữ liệu sai hoặc lệch dòng.
- **Callback Idle Guard**: Tuyệt đối không gọi `cb.PerformCallback()` khi `cb.InCallback() === true`. Luôn chờ component rảnh rỗi trước khi gửi lệnh mới.
- **EndCallback Verification**: Khi xử lý callback bất đồng bộ:
  1. Phải đăng ký sự kiện `EndCallback`.
  2. Phải kiểm tra ID dữ liệu trả về (`txtID.GetText() === targetQid`).
  3. Không đọc DOM khi callback bị timeout mà phải kích hoạt cơ chế Retry với độ trễ an toàn.
- **Client Control API over Raw DOM**: Ưu tiên đọc dữ liệu thông qua DevExpress Client API (`aspxGetControlCollection().Get(...)`, `GetItemCount()`, `GetItem(i).text`) để đảm bảo dữ liệu sạch, đã chuẩn hóa và không trùng lặp.

## 2. Audit Spreadsheet Parsing Guardrails
- **Header vs Question Distinction**: Các hàng tiêu đề mục lớn (như `1. Material storage: The ability to...`) không được tính là câu hỏi kiểm toán.
- **Question Criteria**: Một dòng chỉ được coi là câu hỏi khi có điểm số kiểm toán hợp lệ (`points`) hoặc có mã phân cấp câu hỏi rõ ràng (ví dụ: `1-1`, `2-1`, `12.1`).

## 3. Strict 1-to-1 Data Reconciliation
- **No Duplicate Mapping**: Mỗi dòng trên Web chỉ được map với duy nhất 1 dòng tương ứng trong file chuẩn (Excel). Luôn duy trì tập hợp `used_indices` để khóa các dòng đã được đối chiếu.
- **Positional Priority**: Khi số lượng câu hỏi và thứ tự câu hỏi giữa Web và Form tương đồng, ưu tiên đối chiếu theo thứ tự dòng kết hợp xác thực tương đồng nội dung.
