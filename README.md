# GitHub_API_Client_Excel

Module VBA compact — chỉ phục vụ ghi danh sách Issues từ GitHub ra bảng Excel (`ListObject`).

## Cài đặt

### Files cần import

| File | Mô tả |
|------|--------|
| `src/GitHub_API_Client_Excel.bas` | Module chính |
| `src/JsonConverter.bas` | Parse JSON (Tim Hall, MIT) |

### Bước thực hiện (Windows)

1. Mở Excel → **Alt + F11** → VBA Editor
2. **File → Import File...** → chọn `GitHub_API_Client_Excel.bas`
3. Lặp lại bước 2 cho `JsonConverter.bas`
4. **Tools → References...** → tick **Microsoft Scripting Runtime** → OK

### Bước thực hiện (macOS — Excel 365 for Mac)

Module hỗ trợ cả Windows và macOS qua biên dịch có điều kiện (`#If Mac Then`).

1. Mở Excel → **Tùy chọn (Option) + F11** → VBA Editor
2. Import `GitHub_API_Client_Excel.bas` và `JsonConverter.bas`
3. Import thêm **`Dictionary.cls`** (bản Mac của VBA-tools, tải tại [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)) — vì Mac không có `Microsoft Scripting Runtime`
4. Lần chạy đầu tiên, Excel sẽ hỏi quyền chạy lệnh shell (`curl`) → bấm **OK / Allow**

| Môi trường | HTTP | JSON parse |
|-----------|------|------------|
| Windows | `MSXML2.ServerXMLHTTP.6.0` (sẵn có) | `Microsoft Scripting Runtime` |
| macOS | `curl` qua `MacScript` (tự động) | import `Dictionary.cls` |

## Tạo Token

1. Truy cập https://github.com/settings/tokens (tab **Tokens (classic)**)
2. **Generate new token (classic)**
3. Tick scope:
   - `read:project` — để lấy project fields (Status, Priority, v.v.)
   - `repo` — nếu cần đọc repo private
4. **Generate token** → copy token (bắt đầu bằng `ghp_...`)

> Nếu chỉ đọc repo **public** và không cần project fields, tạo Fine-grained token không cần tick gì cũng được.

## Sử dụng

### Cơ bản — Ghi tất cả issues ra bảng

```vb
Sub Main()
    SetGitHubToken "ghp_your_token_here"
    WriteRepoIssuesTable "octocat", "Hello-World"
End Sub
```

Kết quả: tạo một sheet mới tên `octocat_Hello-World` với bảng 11 cột mặc định + project fields (nếu có). Nếu sheet đã tồn tại, nội dung sẽ được ghi đè.

### Lọc theo trạng thái

```vb
Sub ChiLayOpen()
    SetGitHubToken "ghp_your_token_here"
    
    ' Tham số thứ 3 là states: "ALL" (mặc định), "OPEN", hoặc "CLOSED"
    WriteRepoIssuesTable "owner", "repo", "OPEN"
End Sub
```

### Chọn cột cụ thể

```vb
Sub ChonCot()
    SetGitHubToken "ghp_your_token_here"
    
    ' states đứng trước columns
    WriteRepoIssuesTable "owner", "repo", "OPEN", _
        Array("Number", "Title", "State", "Author", "Labels", "[My Project] Status")
End Sub
```

### Lấy nhiều hơn 100 issues (tự động pagination)

```vb
Sub LayNhieu()
    SetGitHubToken "ghp_your_token_here"
    
    SetMaxIssues 500   ' Tự động gọi nhiều lần (mỗi lần 100) cho đến khi đủ
    WriteRepoIssuesTable "owner", "repo"
End Sub
```

## API Reference

| Thủ tục | Mô tả |
|---------|--------|
| `SetGitHubToken token` | Cấu hình Personal Access Token |
| `SetMaxIssues n` | Đặt số issue tối đa (≥ 1, mặc định 100). Nếu > 100 sẽ tự pagination |
| `WriteRepoIssuesTable id, repo, [states], [columns]` | Ghi bảng issues ra sheet mới |

### WriteRepoIssuesTable

```vb
Public Sub WriteRepoIssuesTable(ByVal githubId As String, ByVal repoName As String, Optional ByVal states As String = "ALL", Optional ByVal columns As Variant)
```

**Tham số:**
- `githubId` — GitHub username hoặc tổ chức
- `repoName` — Tên repository
- `states` (Optional) — `"ALL"` (mặc định), `"OPEN"`, hoặc `"CLOSED"`
- `columns` (Optional) — Mảng tên cột cần xuất; nếu bỏ qua thì xuất tất cả

> Dữ liệu được ghi vào một **sheet mới** tên `{githubId}_{repoName}` (ký tự không hợp lệ thay bằng `_`, tối đa 31 ký tự). Nếu sheet đã tồn tại, nội dung cũ bị xóa và ghi lại — **không bao giờ động đến các sheet khác của bạn**.

### SetMaxIssues

```vb
Public Sub SetMaxIssues(ByVal maxIssues As Long)
```

- Mặc định: `100`
- Giá trị `> 100` → module tự động gọi GitHub API nhiều lần (mỗi lần 100) bằng cursor pagination cho đến khi đủ số lượng hoặc hết issues
- Giá trị `< 1` → phát sinh lỗi runtime

## Tham số `columns`

| Loại | Format | Ví dụ |
|------|--------|-------|
| Issue field | Tên trường (case-insensitive) | `"Number"`, `"Title"`, `"State"` |
| Project field | `"[ProjectTitle] FieldName"` | `"[Sprint Board] Status"` |

11 issue fields hợp lệ: `Number`, `Title`, `Url`, `State`, `Author`, `CreatedAt`, `UpdatedAt`, `ClosedAt`, `Labels`, `Assignees`, `Milestone`

Nếu không truyền `columns` → xuất tất cả 11 + project fields phát hiện được (sorted A-Z).

> **Lưu ý:** Tên cột không khớp issue field và không đúng format `"[...] ..."` sẽ bị bỏ qua.

## Trạng thái Issue

GitHub Issues chỉ có 2 state native: **`OPEN`** và **`CLOSED`**.

Nếu issue thuộc một GitHub Project v2, các field tùy chỉnh (Status, Priority, v.v.) nằm trong cột project field dạng `"[ProjectTitle] FieldName"`, không phải cột `State`.

## Cấu trúc ProjectFields

Mỗi issue có dữ liệu project fields được serialize thành JSON nội bộ:

```json
[{"project":"Sprint Board","field":"Status","value":"In Progress"},{"project":"Sprint Board","field":"Priority","value":"High"}]
```

Khi xuất bảng, mỗi project field thành một cột riêng (header dạng `[Sprint Board] Status`).

## Mã lỗi (trong ErrorMessage khi Err.Raise)

| ErrorCode | Ý nghĩa |
|-----------|---------|
| `MISSING_TOKEN` | Chưa gọi `SetGitHubToken` |
| `INVALID_GITHUB_ID` | githubId rỗng/sai ký tự/quá 39 ký tự |
| `INVALID_REPO_NAME` | repoName rỗng/sai ký tự/quá 100 ký tự |
| `INVALID_STATES` | states phải là OPEN, CLOSED, hoặc ALL |
| `NOT_FOUND` | Repo không tồn tại |
| `FORBIDDEN` | Không đủ quyền |
| `UNAUTHORIZED` | Token hết hạn hoặc sai |
| `RATE_LIMITED` | Vượt giới hạn API (5000 điểm/giờ) |
| `SERVER_ERROR` | GitHub lỗi (5xx) |
| `HTTP_ERROR` | Lỗi HTTP khác |
| `NETWORK_ERROR` | Không có mạng/DNS/SSL |
| `PARSE_ERROR` | Response không phải JSON hợp lệ |
| `GRAPHQL_ERROR` | Lỗi GraphQL khác |

## Lỗi runtime (Err.Raise)

| Số lỗi | Khi nào |
|---------|---------|
| `vbObjectError + 1001` | `WriteRepoIssuesTable` — API thất bại (ErrorCode + ErrorMessage trong Err.Description) |
| `vbObjectError + 1003` | `SetMaxIssues` — giá trị < 1 |

Bắt lỗi bằng `On Error`:

```vb
Sub SafeRun()
    On Error GoTo Failed
    SetGitHubToken "ghp_your_token_here"
    WriteRepoIssuesTable "owner", "repo", "OPEN"
    Exit Sub
Failed:
    MsgBox "Lỗi: " & Err.Description
End Sub
```

## Yêu cầu

- Excel 2007+ trên Windows
- Kết nối Internet
- Reference: Microsoft Scripting Runtime
- GitHub Personal Access Token
