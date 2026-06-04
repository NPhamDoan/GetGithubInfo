# GitHub Issues qua Office Scripts (TypeScript)

Giải pháp chạy được **Excel Windows + Mac + Web**, **không cần VBA, không cần curl**. Dùng Office Scripts (`fetch` built-in).

## Yêu cầu

- **Microsoft 365 business/enterprise** (có tab **Automate** trên ribbon)
- File lưu trên **OneDrive for Business** hoặc **SharePoint**
- ❌ Không chạy với Office mua lẻ (2019/2021 standalone) hoặc tài khoản cá nhân

> Nếu không có tab Automate → dùng Power Query (`README_PowerQuery.md`) thay thế.

## Cài đặt

1. Mở Excel → tab **Automate** → **New Script**
2. Xóa nội dung mẫu, dán toàn bộ `src/GitHubIssues.ts`
3. Sửa phần **CONFIG** ở đầu hàm `main`:
   ```ts
   const owner = "octocat";
   const repo = "Hello-World";
   const token = "ghp_...";          // PAT scope read:project
   const states = "ALL";             // "ALL" | "OPEN" | "CLOSED"
   const maxIssues = 100;            // > 100 tự pagination
   const columns: string[] = [];     // [] = tất cả cột
   ```
4. Bấm **Run**

Kết quả: tạo sheet mới tên `{owner}_{repo}` + Table chứa issues. Chạy lại sẽ xóa sheet cũ và tạo mới.

## Tạo Token

1. https://github.com/settings/tokens → **Tokens (classic)**
2. **Generate new token (classic)** → tick `read:project` (và `repo` nếu private)
3. Copy `ghp_...` vào biến `token`

## Chọn cột (giống tham số columns trong VBA)

Để `columns = []` → xuất tất cả 11 issue fields + project fields tìm thấy.

Hoặc liệt kê cụ thể (theo thứ tự):
```ts
const columns: string[] = [
  "Number", "Title", "State", "Author", "Labels",
  "[Sprint Board] Status", "[Sprint Board] Priority"
];
```

- Issue fields: `Number, Title, Url, State, Author, CreatedAt, UpdatedAt, ClosedAt, Labels, Assignees, Milestone`
- Project fields: `"[ProjectTitle] FieldName"`
- Cột không tồn tại → để trống

## Lọc theo trạng thái

```ts
const states = "OPEN";   // hoặc "CLOSED" / "ALL"
```

## Lấy hơn 100 issues

```ts
const maxIssues = 500;   // tự gọi nhiều lần (mỗi lần 100)
```

## Tạo nút bấm (tùy chọn)

Sau khi lưu script: tab **Automate → script của bạn → ... → Add button**. Nút xuất hiện trên sheet, bấm là chạy.

## Chia sẻ cho nhiều người

- Script lưu trong OneDrive/SharePoint của bạn → **Share** với nhóm
- Mỗi người chạy bằng tài khoản của họ
- **Token:** nên để mỗi người tự điền token của họ vào CONFIG (đừng share token cá nhân)

## So sánh với các giải pháp khác

| | VBA | Power Query | Office Scripts |
|---|---|---|---|
| Windows | ✅ | ✅ | ✅ |
| macOS | ⚠️ cần curl | ✅ | ✅ |
| Excel Web | ❌ | ⚠️ | ✅ |
| Cần curl | Mac cần | Không | Không |
| Yêu cầu | Office bất kỳ | Office bất kỳ | M365 business |
| Chạy bằng | Macro/UDF | Refresh | Run / nút bấm |

## Lưu ý kỹ thuật

- `fetch()` là built-in trong Office Scripts (không phải Node.js fetch — chạy trong sandbox Excel)
- Ngày tháng trả về dạng chuỗi ISO 8601; muốn format Date thì xử lý thêm hoặc đổi định dạng cột sau
- Office Scripts có giới hạn thời gian chạy (~5 phút) — đủ cho vài trăm issues
