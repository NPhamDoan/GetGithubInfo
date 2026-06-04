# VBA GitHub Repo Info

Mô-đun VBA cho phép lấy thông tin repository GitHub và danh sách Issues kèm project fields (GitHub Projects v2) trực tiếp từ Microsoft Office (Excel, Word, Access) thông qua GitHub GraphQL API.

## Tính năng chính

- **GetRepoInfo** — Lấy thông tin tổng quan repository (mô tả, ngôn ngữ, stars, forks, issues, v.v.)
- **GetRepoField** — Hàm Excel UDF lấy từng trường thông tin repo trong ô bảng tính
- **GetRepoIssues** — Lấy danh sách Issues kèm thông tin project fields (trạng thái, priority, custom fields từ GitHub Projects v2)
- **GetRepoIssueField** — Hàm Excel UDF lấy từng trường thông tin issue trong ô bảng tính
- **WriteRepoIssuesTable** — Ghi danh sách issues ra bảng Excel (ListObject) với cột tùy chọn

## Hướng dẫn nhập module vào Excel/Word/Access

### Các file cần nhập

| File | Mô tả |
|------|--------|
| `src/GitHub_API_Client.bas` | Module chính chứa toàn bộ API |
| `src/JsonConverter.bas` | Module parse JSON (Tim Hall, MIT License) |

### Bước thực hiện

1. Mở ứng dụng Office (Excel, Word hoặc Access)
2. Nhấn **Alt + F11** để mở VBA Editor
3. Trong VBA Editor, chọn menu **File → Import File...**
4. Chọn file `GitHub_API_Client.bas` và nhấn **Open**
5. Lặp lại bước 3-4 để nhập file `JsonConverter.bas`
6. Đóng VBA Editor (Ctrl + Q hoặc nút X)

> **Lưu ý:** Cả hai module đều sử dụng late binding (`CreateObject`) nên **không cần** thêm reference thủ công nào trên Windows. Module hoạt động trên cả Office 32-bit và 64-bit.

### Chạy trên macOS (Excel 365 for Mac)

Module hỗ trợ **cả Windows và macOS** thông qua biên dịch có điều kiện (`#If Mac Then`):

- **HTTP:** Trên Mac, `MSXML2.ServerXMLHTTP.6.0` không tồn tại → module tự động dùng `curl` qua `MacScript`. Cần cấp quyền cho Excel chạy lệnh shell (Excel sẽ hỏi lần đầu).
- **JSON / Dictionary:** Trên Mac không có `Microsoft Scripting Runtime`. Bạn cần import thêm file **`Dictionary.cls`** (bản Mac của VBA-tools) để `JsonConverter` hoạt động. Tải tại [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary).
- Module gốc đã thay `Scripting.Dictionary` bằng `Collection` (built-in) ở phần xử lý nội bộ nên chỉ còn `JsonConverter` cần `Dictionary.cls` trên Mac.

| Môi trường | HTTP | JSON parse |
|-----------|------|------------|
| Windows | `MSXML2.ServerXMLHTTP.6.0` (sẵn có) | `Scripting.Dictionary` (sẵn có) |
| macOS | `curl` qua `MacScript` | cần import `Dictionary.cls` |

### Tham chiếu thư viện

Module sử dụng các thư viện có sẵn, truy cập qua `CreateObject` (late binding):

- **Windows:** `MSXML2.ServerXMLHTTP.6.0`, `Scripting.Dictionary`
- **macOS:** `curl` (có sẵn trong macOS), `Dictionary.cls` (import thủ công)

## Tạo Personal Access Token trên GitHub

Để sử dụng module, bạn cần tạo Personal Access Token (PAT) trên GitHub:

1. Truy cập [GitHub Settings → Developer settings → Personal access tokens → Fine-grained tokens](https://github.com/settings/tokens?type=beta)
2. Nhấn **Generate new token**
3. Đặt tên token (ví dụ: "VBA GitHub Client")
4. Chọn thời hạn (Expiration)
5. Trong mục **Repository permissions**, cấp quyền:
   - `Contents` → Read-only (để đọc thông tin repo)
6. Trong mục **Account permissions** hoặc **Organization permissions**, cấp scope `read:project` để sử dụng thông tin project fields trong `GetRepoIssues` và trường `RepoInfo.ProjectsCount`
   - Nếu cần quyền ghi project, sử dụng scope `project`
7. Nhấn **Generate token** và sao chép token

> **Quan trọng:** Token cần có scope `read:project` (hoặc `project` cho quyền ghi) để truy cập thông tin project fields trong `GetRepoIssues` (trường `ProjectFields` của mỗi issue) và trường `ProjectsCount` trong `RepoInfo`. Nếu không có scope này, các trường liên quan đến project sẽ trống hoặc trả lỗi FORBIDDEN.

## Cấu hình Token

Sử dụng thủ tục `SetGitHubToken` để cấu hình token trong phiên VBA hiện tại:

```vb
Sub CauHinhToken()
    ' Cấu hình Personal Access Token
    SetGitHubToken "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
End Sub
```

> **Lưu ý:** Token chỉ tồn tại trong phiên VBA hiện tại. Khi đóng và mở lại file, bạn cần gọi `SetGitHubToken` lại. Nếu chưa cấu hình token, mọi hàm sẽ trả về lỗi `MISSING_TOKEN`.

## Ví dụ sử dụng

### (a) Gọi `GetRepoInfo` từ Sub VBA

```vb
Sub LayThongTinRepo()
    ' Cấu hình token trước khi gọi
    SetGitHubToken "ghp_your_token_here"
    
    ' Gọi hàm lấy thông tin repository
    Dim info As RepoInfo
    info = GetRepoInfo("octocat", "Hello-World")
    
    ' Kiểm tra kết quả
    If info.Success Then
        Debug.Print "Repo: " & info.NameWithOwner
        Debug.Print "Mô tả: " & info.Description
        Debug.Print "Ngôn ngữ: " & info.PrimaryLanguage
        Debug.Print "Stars: " & info.Stars
        Debug.Print "Forks: " & info.Forks
        Debug.Print "Open Issues: " & info.OpenIssues
        Debug.Print "Default Branch: " & info.DefaultBranch
        Debug.Print "URL: " & info.Url
        Debug.Print "Private: " & info.IsPrivate
        Debug.Print "Archived: " & info.IsArchived
        Debug.Print "Disk Usage (KB): " & info.DiskUsageKB
        Debug.Print "Projects Count: " & info.ProjectsCount
        Debug.Print "Created: " & info.CreatedAt
        Debug.Print "Updated: " & info.UpdatedAt
    Else
        Debug.Print "Lỗi: " & info.ErrorCode & " - " & info.ErrorMessage
    End If
End Sub
```

### (b) Gọi `GetRepoField` từ ô Excel

Trong ô Excel, nhập công thức:

```
=GetRepoField("octocat", "Hello-World", "Stars")
```

Các giá trị `fieldName` hợp lệ:

| fieldName | Kiểu trả về | Mô tả |
|-----------|-------------|--------|
| `"NameWithOwner"` | String | Tên đầy đủ (owner/repo) |
| `"Description"` | String | Mô tả repository |
| `"PrimaryLanguage"` | String | Ngôn ngữ chính |
| `"Stars"` | Long | Số stars |
| `"Forks"` | Long | Số forks |
| `"OpenIssues"` | Long | Số issues đang mở |
| `"DefaultBranch"` | String | Nhánh mặc định |
| `"Url"` | String | URL repository |
| `"CreatedAt"` | Date | Ngày tạo |
| `"UpdatedAt"` | Date | Ngày cập nhật |
| `"IsPrivate"` | Boolean | Repository có private không |
| `"IsFork"` | Boolean | Repository có phải fork không |
| `"IsArchived"` | Boolean | Repository có bị archive không |
| `"DiskUsageKB"` | Long | Dung lượng đĩa (KB) |
| `"ProjectsCount"` | Long | Số lượng projects |

> Nếu `fieldName` không hợp lệ, hàm trả về `#VALUE!`. Nếu API gặp lỗi, hàm trả về `#N/A`.

### (c) Gọi `GetRepoIssues` từ Sub VBA

```vb
Sub LayDanhSachIssues()
    SetGitHubToken "ghp_your_token_here"
    
    ' Tùy chọn: thay đổi số issue tối đa (mặc định 100; > 100 sẽ tự pagination)
    SetMaxIssues 50
    
    ' Gọi hàm lấy danh sách issues (states: "ALL", "OPEN", hoặc "CLOSED")
    Dim result As RepoIssuesResult
    GetRepoIssues "octocat", "Hello-World", result, "OPEN"
    
    If result.Success Then
        Debug.Print "Tổng số issues: " & result.TotalCount
        
        Dim i As Long
        Dim issue As RepoIssue
        For i = 1 To result.TotalCount
            issue = GetIssueItem(i)
            Debug.Print "---"
            Debug.Print "Issue #" & issue.Number & ": " & issue.Title
            Debug.Print "  URL: " & issue.Url
            Debug.Print "  State: " & issue.State
            Debug.Print "  Author: " & issue.Author
            Debug.Print "  Labels: " & issue.Labels
            Debug.Print "  Assignees: " & issue.Assignees
            Debug.Print "  Milestone: " & issue.Milestone
            Debug.Print "  Created: " & issue.CreatedAt
            Debug.Print "  Updated: " & issue.UpdatedAt
            Debug.Print "  Closed: " & issue.ClosedAt
            Debug.Print "  ProjectFields: " & issue.ProjectFields
        Next i
    Else
        Debug.Print "Lỗi: " & result.ErrorCode & " - " & result.ErrorMessage
    End If
End Sub
```

### (d) Gọi `GetRepoIssueField` từ ô Excel

Trong ô Excel, nhập công thức:

```
=GetRepoIssueField("octocat", "Hello-World", 1, "Title")
```

Tham số:
- `githubId` — GitHub username hoặc tổ chức
- `repoName` — Tên repository
- `issueIndex` — Số thứ tự issue trong kết quả (bắt đầu từ 1)
- `fieldName` — Tên trường cần lấy

Các giá trị `fieldName` hợp lệ cho `GetRepoIssueField`:

| fieldName | Kiểu trả về | Mô tả |
|-----------|-------------|--------|
| `"Number"` | Long | Số hiệu issue |
| `"Title"` | String | Tiêu đề issue |
| `"Url"` | String | URL issue trên GitHub |
| `"State"` | String | Trạng thái (`"OPEN"` hoặc `"CLOSED"`) |
| `"Author"` | String | Login của tác giả |
| `"CreatedAt"` | Date | Ngày tạo |
| `"UpdatedAt"` | Date | Ngày cập nhật |
| `"ClosedAt"` | Date | Ngày đóng (rỗng nếu chưa đóng) |
| `"Labels"` | String | Danh sách labels (phân cách bởi ", ") |
| `"Assignees"` | String | Danh sách assignees (phân cách bởi ", ") |
| `"Milestone"` | String | Tên milestone |
| `"ProjectFields"` | String | JSON array chứa project field values |

> Nếu `fieldName` không hợp lệ → `#VALUE!`. Nếu `issueIndex` ngoài phạm vi hoặc API lỗi → `#N/A`.

### (e) Gọi `WriteRepoIssuesTable` từ Sub VBA

```vb
Sub GhiIssuesRaBang()
    SetGitHubToken "ghp_your_token_here"
    
    ' Tạo sheet mới tên "octocat_Hello-World" và ghi bảng issues vào đó
    ' Mặc định xuất tất cả 11 cột issue + project fields tìm thấy (states "ALL")
    WriteRepoIssuesTable "octocat", "Hello-World"
    
    ' Lọc theo trạng thái (tham số thứ 3): "ALL", "OPEN", hoặc "CLOSED"
    WriteRepoIssuesTable "octocat", "Hello-World", "OPEN"
    
    ' Chỉ định cột cụ thể (states đứng trước columns):
    WriteRepoIssuesTable "octocat", "Hello-World", "OPEN", _
        Array("Number", "Title", "State", "Author", "Labels", "[Sprint Board] Status", "[Sprint Board] Priority")
    
    ' Kết quả: tạo sheet mới + ListObject (Table) với các cột đã chọn
    ' - Sheet tên "octocat_Hello-World" (ghi đè nếu đã tồn tại, không động sheet khác)
    ' - Cột CreatedAt, UpdatedAt, ClosedAt định dạng "yyyy-mm-dd hh:mm:ss"
    ' - ClosedAt hiển thị chuỗi rỗng nếu issue chưa đóng
    ' - Bảng được AutoFit để vừa nội dung
    ' - Tên bảng: tblRepoIssues_octocat_Hello_World
End Sub
```

**Giải thích tham số `columns`:**

- Nếu không truyền `columns` → xuất tất cả 11 issue fields mặc định + tất cả project fields phát hiện được (sắp xếp alphabet)
- 11 issue fields mặc định: `Number`, `Title`, `Url`, `State`, `Author`, `CreatedAt`, `UpdatedAt`, `ClosedAt`, `Labels`, `Assignees`, `Milestone`
- Project fields sử dụng format: `"[ProjectTitle] FieldName"` (ví dụ: `"[Sprint Board] Status"`)
- Tên issue field so khớp case-insensitive
- Tên FieldName trong project field so khớp case-insensitive; ProjectTitle giữ nguyên case

> **Lưu ý:** Nếu ô đích đã thuộc một `ListObject` hiện có, thủ tục sẽ xóa dữ liệu cũ và ghi đè dữ liệu mới mà không tạo bảng mới. Nếu token chưa cấu hình hoặc API gặp lỗi, thủ tục sẽ phát sinh lỗi runtime (`Err.Raise`) để người gọi xử lý bằng `On Error`.

## Cấu trúc ProjectFields

Mỗi issue có trường `ProjectFields` chứa JSON array serialize các giá trị project field:

```json
[
  {"project": "Sprint Board", "field": "Status", "value": "In Progress"},
  {"project": "Sprint Board", "field": "Priority", "value": "High"},
  {"project": "Roadmap", "field": "Quarter", "value": "Q1 2025"}
]
```

- `project` — Tên GitHub Project v2 mà issue thuộc
- `field` — Tên field trong project
- `value` — Giá trị field (string representation)
- Chuỗi `"[]"` khi issue không thuộc project nào

### Parse ProjectFields trong VBA

```vb
Sub ParseProjectFields()
    Dim jsonStr As String
    Dim issue As RepoIssue
    issue = GetIssueItem(1)  ' Lấy issue đầu tiên từ kết quả GetRepoIssues gần nhất
    jsonStr = issue.ProjectFields  ' Ví dụ: [{"project":"Sprint Board","field":"Status","value":"In Progress"}]
    
    ' Sử dụng JsonConverter để parse
    Dim fields As Object
    Set fields = JsonConverter.ParseJson(jsonStr)
    
    ' Duyệt từng field value
    Dim item As Object
    For Each item In fields
        Debug.Print "Project: " & item("project")
        Debug.Print "Field: " & item("field")
        Debug.Print "Value: " & item("value")
    Next item
End Sub
```

## Bảng mã lỗi (ErrorCode)

Bảng dưới đây liệt kê đầy đủ các giá trị `ErrorCode` có thể xuất hiện trong `RepoInfo.ErrorCode` hoặc `RepoIssuesResult.ErrorCode`:

| # | ErrorCode | Ý nghĩa |
|---|-----------|---------|
| 1 | `""` (chuỗi rỗng) | Thành công — không có lỗi |
| 2 | `MISSING_TOKEN` | Token chưa cấu hình. Gọi `SetGitHubToken` trước khi sử dụng |
| 3 | `INVALID_GITHUB_ID` | Tham số `githubId` không hợp lệ: rỗng, chứa whitespace, chứa ký tự ngoài `[A-Za-z0-9-]`, hoặc dài hơn 39 ký tự |
| 4 | `INVALID_REPO_NAME` | Tham số `repoName` không hợp lệ: rỗng, chứa whitespace, chứa ký tự ngoài `[A-Za-z0-9._-]`, hoặc dài hơn 100 ký tự |
| 5 | `INVALID_STATES` | Tham số `states` không hợp lệ: phải là `"OPEN"`, `"CLOSED"` hoặc `"ALL"` (case-insensitive) |
| 6 | `NOT_FOUND` | Repository không tồn tại hoặc không có quyền truy cập |
| 7 | `FORBIDDEN` | Bị từ chối truy cập (token không đủ quyền cho tài nguyên này) |
| 8 | `UNAUTHORIZED` | Token không hợp lệ hoặc đã hết hạn. Tạo token mới và gọi `SetGitHubToken` lại |
| 9 | `RATE_LIMITED` | Vượt quá giới hạn tốc độ truy cập GitHub API (5000 điểm/giờ). Chờ và thử lại sau |
| 10 | `SERVER_ERROR` | Lỗi máy chủ GitHub (HTTP 5xx). Thử lại sau vài phút |
| 11 | `HTTP_ERROR` | Lỗi HTTP không xác định (mã trạng thái nằm ngoài các trường hợp đã xử lý) |
| 12 | `NETWORK_ERROR` | Lỗi mạng — không có kết nối Internet, DNS không phân giải, hoặc lỗi SSL |
| 13 | `TIMEOUT` | Yêu cầu vượt quá thời gian chờ 30 giây |
| 14 | `PARSE_ERROR` | Phản hồi từ GitHub không phải JSON hợp lệ hoặc thiếu trường bắt buộc |
| 15 | `GRAPHQL_ERROR` | Lỗi GraphQL không thuộc các loại đã phân loại ở trên |

## Cấu hình nâng cao

### Thay đổi số issue tối đa (SetMaxIssues)

Mặc định, `GetRepoIssues` lấy tối đa **100** issues mỗi lần gọi. Bạn có thể đặt giá trị lớn hơn — khi đó module tự động gọi GitHub API nhiều lần (mỗi lần 100) bằng cursor pagination cho đến khi đủ số lượng hoặc hết issues:

```vb
SetMaxIssues 50    ' Lấy tối đa 50 issues (1 lần gọi)
SetMaxIssues 500   ' Lấy tối đa 500 issues (tự gọi 5 lần x 100)
```

Chữ ký:

```vb
Public Sub SetMaxIssues(ByVal maxIssues As Long)
```

- Giá trị mặc định: `100`
- Phạm vi hợp lệ: `>= 1` (giá trị > 100 kích hoạt pagination tự động)
- Nếu giá trị < 1, thủ tục sẽ phát sinh lỗi (`Err.Raise`) và giữ nguyên giá trị cũ

## Chữ ký API

### GetRepoIssues

```vb
Public Sub GetRepoIssues(ByVal githubId As String, ByVal repoName As String, ByRef outResult As RepoIssuesResult, Optional ByVal states As String = "ALL")
```

**Tham số:**
- `githubId` — GitHub username hoặc tên tổ chức (tối đa 39 ký tự, chỉ `[A-Za-z0-9-]`)
- `repoName` — Tên repository (tối đa 100 ký tự, chỉ `[A-Za-z0-9._-]`)
- `outResult` (ByRef) — Biến `RepoIssuesResult` nhận kết quả trả về (gồm `TotalCount`, `Success`, `ErrorCode`, `ErrorMessage`)
- `states` (Optional) — Bộ lọc trạng thái: `"ALL"` (mặc định), `"OPEN"`, hoặc `"CLOSED"` (case-insensitive)

**Truy cập từng issue:** Dùng `GetIssueItem(index)` (1-based) sau khi gọi `GetRepoIssues`

### GetRepoIssueField

```vb
Public Function GetRepoIssueField(ByVal githubId As String, ByVal repoName As String, ByVal issueIndex As Long, ByVal fieldName As String) As Variant
```

**Tham số:**
- `githubId` — GitHub username hoặc tổ chức
- `repoName` — Tên repository
- `issueIndex` — Số thứ tự issue (bắt đầu từ 1)
- `fieldName` — Tên trường cần lấy (case-insensitive)

**Trả về:** Giá trị trường tương ứng, `#VALUE!` nếu fieldName không hợp lệ, `#N/A` nếu lỗi hoặc index ngoài phạm vi

### GetRepoIssuesJson

```vb
Public Function GetRepoIssuesJson(ByVal githubId As String, ByVal repoName As String, Optional ByVal states As String = "ALL") As String
```

**Tham số:**
- `githubId` — GitHub username hoặc tổ chức
- `repoName` — Tên repository
- `states` (Optional) — `"ALL"` (mặc định), `"OPEN"`, hoặc `"CLOSED"`

**Trả về:** Chuỗi JSON `{"success":bool,"totalCount":N,"errorCode":"","errorMessage":"","issues":[...]}`. Dùng khi muốn xử lý dữ liệu trong code thay vì ghi ra sheet.

Ví dụ:

```vb
Sub LayJson()
    SetGitHubToken "ghp_your_token_here"
    Dim json As String
    json = GetRepoIssuesJson("octocat", "Hello-World", "OPEN")
    Debug.Print json

    ' Parse lại để duyệt
    Dim data As Object
    Set data = JsonConverter.ParseJson(json)
    If data("success") Then
        Dim iss As Object
        For Each iss In data("issues")
            Debug.Print iss("number") & ": " & iss("title")
        Next iss
    End If
End Sub
```

Mỗi issue trong `issues`: `number, title, url, state, author, createdAt, updatedAt, closedAt, labels, assignees, milestone, projectFields` (projectFields là JSON array thật).

### WriteRepoIssuesTable

```vb
Public Sub WriteRepoIssuesTable(ByVal githubId As String, ByVal repoName As String, Optional ByVal states As String = "ALL", Optional ByVal columns As Variant)
```

**Tham số:**
- `githubId` — GitHub username hoặc tổ chức
- `repoName` — Tên repository
- `states` (Optional) — Bộ lọc trạng thái: `"ALL"` (mặc định), `"OPEN"`, hoặc `"CLOSED"`
- `columns` (Optional) — Mảng tên cột cần xuất; nếu không truyền thì xuất tất cả

> Dữ liệu được ghi vào một **sheet mới** tên `{githubId}_{repoName}`. Nếu sheet đã tồn tại, nội dung cũ bị xóa và ghi lại — không động đến các sheet khác.

**Lỗi runtime:**
- `vbObjectError + 1001` — GetRepoIssues thất bại (ErrorCode + ErrorMessage trong Err.Description)

### SetMaxIssues

```vb
Public Sub SetMaxIssues(ByVal maxIssues As Long)
```

**Tham số:**
- `maxIssues` — Số issue tối đa mỗi lần gọi, mặc định `100`. Giá trị > 100 sẽ tự động pagination (gọi nhiều lần x 100)

**Lỗi runtime:** `vbObjectError + 1003` nếu giá trị < 1

## Yêu cầu hệ thống

- Microsoft Office 2007 trở lên (Excel, Word, hoặc Access)
- **Windows:** dùng `MSXML2.ServerXMLHTTP.6.0` + `Scripting.Dictionary` (sẵn có)
- **macOS (Excel 365):** dùng `curl` (sẵn có) + import `Dictionary.cls` cho `JsonConverter`
- Kết nối Internet
- Personal Access Token từ GitHub (thêm scope `read:project` nếu cần project fields)

## Giấy phép

- `GitHub_API_Client.bas` — MIT License
- `JsonConverter.bas` — MIT License (Tim Hall, [VBA-JSON](https://github.com/VBA-tools/VBA-JSON))
