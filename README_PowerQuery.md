# GitHub Issues qua Power Query (Windows + macOS)

Giải pháp thay thế VBA, **chạy được cả Windows và Mac**, **không cần curl, không cần VBA, không cần reference**. Dùng Power Query (Get & Transform) tích hợp sẵn trong Excel.

## Khi nào dùng cái này thay vì VBA?

| | VBA (`.bas`) | Power Query |
|---|---|---|
| Windows | ✅ | ✅ |
| macOS | ⚠️ cần curl (có thể bị chặn) | ✅ không cần shell |
| Cần cài thêm | Import `.bas` | Không, có sẵn trong Excel |
| Cách dùng | Gọi macro / UDF | Bấm **Refresh** để cập nhật bảng |

> Nếu máy công ty chặn shell/curl trên Mac → **dùng Power Query**.

## Cài đặt (làm 1 lần cho mỗi workbook)

1. Mở Excel → tab **Data** → **Get Data** → **From Other Sources** → **Blank Query**
   (Mac: **Data → Get Data → Launch Power Query Editor**)
2. Trong Power Query Editor → **Home → Advanced Editor**
3. Xóa hết nội dung mặc định, dán toàn bộ nội dung file `src/GitHubIssues.pq`
4. Sửa phần **CẤU HÌNH** ở đầu query:
   ```
   Owner     = "octocat",
   Repo      = "Hello-World",
   States    = "ALL",          // "ALL" | "OPEN" | "CLOSED"
   MaxIssues = 100,            // > 100 sẽ tự pagination
   Token     = "ghp_...",      // token của bạn
   ```
5. **Done** → **Close & Load**

Kết quả: một bảng (Table) xuất hiện trên sheet mới, có đầy đủ cột issue + các cột project field dạng `[Project] FieldName`.

## Cập nhật dữ liệu

Bấm chuột phải vào bảng → **Refresh**, hoặc **Data → Refresh All**. Query gọi lại API và cập nhật bảng.

## Tạo Token

1. https://github.com/settings/tokens → tab **Tokens (classic)**
2. **Generate new token (classic)** → tick scope:
   - `read:project` — để lấy project fields
   - `repo` — nếu đọc repo private
3. Copy token (`ghp_...`) vào biến `Token` trong query.

## Bảo mật Token (quan trọng)

Token được lưu trong query của workbook. Nếu chia sẻ file, **người khác đọc được token**. Cách an toàn hơn:

**Tách Token thành Parameter:**
1. Power Query Editor → **Manage Parameters → New Parameter**
2. Tên: `Token`, kiểu Text
3. Trong query, đổi dòng `Token = "ghp_..."` thành `Token = TokenParam` (tên parameter)
4. Khi gửi file, xóa giá trị parameter — người nhận tự nhập token của họ

## Cấu hình cho nhiều người dùng

Power Query nằm **trong workbook**, không phải module dùng chung. Để phân phối:

**Cách 1 — File template:** Tạo 1 file `.xlsx` có sẵn query → gửi cho mọi người. Họ chỉ cần đổi Owner/Repo/Token rồi Refresh.

**Cách 2 — Parameter hóa:** Biến `Owner`, `Repo`, `States`, `Token` thành Parameters (như trên) để người dùng đổi qua giao diện **Data → Refresh** mà không cần sửa code M.

## Cột kết quả

- **Issue fields:** Number, Title, Url, State, Author, CreatedAt, UpdatedAt, ClosedAt, Labels, Assignees, Milestone
- **Project fields:** mỗi field thành 1 cột tên `[ProjectTitle] FieldName` (ví dụ `[Sprint Board] Status`)
- Issue không có project field → ô trống (null)

## Lưu ý kỹ thuật

- `Web.Contents` với `Content = Json.FromValue(...)` tự động gửi **POST**
- Header `Authorization: Bearer {token}` được gắn qua tùy chọn `Headers`
- Pagination tự động: query lấy 100 issue/trang đến khi đủ `MaxIssues` hoặc hết
- Ngày tháng (CreatedAt/UpdatedAt/ClosedAt) trả về dạng chuỗi ISO 8601; có thể đổi kiểu cột sang Date trong Power Query nếu cần

## Giới hạn

- Không gọi được như một hàm/macro tùy ý — chỉ refresh bảng
- Mỗi workbook phải có query riêng (không phải module dùng chung như `.bas`)
- Nếu cần ghi vào sheet với logic phức tạp (đặt tên bảng, format đặc thù) → vẫn nên dùng VBA trên Windows
