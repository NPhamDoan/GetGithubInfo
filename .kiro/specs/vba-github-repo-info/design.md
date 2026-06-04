# Design Document

## Overview

`repo-issues-with-project-fields` là một mô-đun chuẩn VBA (`GitHub_API_Client.bas`) cung cấp ba điểm vào chính cho người dùng Microsoft Office (Excel, Word, Access):

1. `GetRepoInfo(githubId, repoName) As RepoInfo` — lấy thông tin tổng quan của một repository GitHub.
2. `GetRepoIssues(githubId, repoName, Optional states) As Variant` — lấy danh sách các Issues kèm thông tin project fields (trạng thái, priority, custom fields từ GitHub Projects v2).
3. `WriteRepoIssuesTable(githubId, repoName, targetRange, Optional columns)` — ghi danh sách issues ra một `ListObject` trên Excel worksheet với cột tùy chọn.

Bên cạnh đó, mô-đun cung cấp các thủ tục cấu hình (`SetGitHubToken`, `SetMaxIssues`) và các hàm phụ phục vụ Excel UDF (`GetRepoField`, `GetRepoIssueField`).

Toàn bộ giao tiếp với GitHub được thực hiện qua một endpoint duy nhất `https://api.github.com/graphql` (HTTP POST, body JSON, GraphQL có tham số hóa). Mô-đun tự gói các tình huống lỗi (xác thực, mạng, timeout, rate limit, GraphQL error, lỗi parse JSON) thành một bảng mã lỗi thống nhất để người gọi không phải xử lý lỗi runtime của HTTP/JSON.

### Mục tiêu thiết kế

- **Đơn giản cho người dùng cuối**: chỉ cần hai tham số bắt buộc (`githubId`, `repoName`) và một lần cấu hình token.
- **An toàn lỗi**: mọi nhánh thực thi đều trả về một `RepoInfo`/`RepoIssuesResult` có trường trạng thái (`Success`, `ErrorCode`, `ErrorMessage`); không ném lỗi runtime ra ngoài (trừ `WriteRepoIssuesTable` theo thiết kế).
- **Đa môi trường**: chạy được trên Excel, Word, Access mà không cần thay đổi mã nguồn, dùng late binding để tránh phụ thuộc reference.
- **Có thể kiểm thử**: tách rõ tầng "pure logic" (validation, build query, parse JSON, mapping lỗi) khỏi tầng "I/O" (HTTP, Excel ListObject) để có thể test bằng property-based testing trên tầng pure logic và bằng mock/integration test trên tầng I/O.

### Phạm vi không xử lý

- Không hỗ trợ pagination cho `issues` (giới hạn cứng tối đa 100 phần tử mỗi lần thông qua `MaxIssues`).
- Không cache kết quả giữa các lần gọi.
- Không hỗ trợ tự động thử lại (retry) khi gặp `RATE_LIMITED`/`SERVER_ERROR`; người dùng tự chờ và gọi lại.
- Không lưu token ra đĩa; token chỉ tồn tại trong phiên VBA.

## Architecture

### Sơ đồ tầng

```
┌─────────────────────────────────────────────────────────────────┐
│                    Môi trường Office VBA                         │
│  ┌──────────────────┐ ┌────────────────┐ ┌───────────────────┐ │
│  │ Excel UDF        │ │ VBA Sub        │ │ Worksheet         │ │
│  │ GetRepoField     │ │ người dùng     │ │ ListObject        │ │
│  │ GetRepoIssueField│ │                │ │ (tblRepoIssues_*) │ │
│  └────────┬─────────┘ └───────┬────────┘ └─────────▲─────────┘ │
└───────────┼────────────────────┼────────────────────┼───────────┘
            │                    │                    │
            ▼                    ▼                    │
┌─────────────────────────────────────────────────────────────────┐
│                   GitHub_API_Client.bas                          │
│                                                                 │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │ Public API                                                  │ │
│  │ GetRepoInfo / GetRepoIssues / WriteRepoIssuesTable /       │ │
│  │ SetGitHubToken / SetMaxIssues / GetRepoField /             │ │
│  │ GetRepoIssueField                                          │ │
│  └────────────────────────────┬───────────────────────────────┘ │
│               │               │               │                 │
│               ▼               ▼               ▼                 │
│  ┌──────────────┐  ┌──────────────────┐  ┌──────────────────┐  │
│  │ Validation   │  │ Query Builder    │  │ Module State     │  │
│  │ Validate     │  │ BuildRepoInfo    │  │ m_Token          │  │
│  │ GithubId     │  │ Query            │  │ m_MaxIssues      │  │
│  │ Validate     │  │ BuildRepoIssues  │  │ m_HttpFactory    │  │
│  │ RepoName     │  │ Query            │  │                  │  │
│  │              │  │ BuildRequestBody │  │                  │  │
│  │              │  │ JsonEscape       │  │                  │  │
│  └──────────────┘  └────────┬─────────┘  └──────────────────┘  │
│                              │                                  │
│                              ▼                                  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ HTTP Transport                                            │   │
│  │ SendGraphQLRequest / CreateHttpClient                     │   │
│  └────────────────────────────┬──────────────────────────────┘  │
│                              │                                  │
│               ┌──────────────┼──────────────┐                   │
│               ▼                             ▼                   │
│  ┌─────────────────────┐       ┌─────────────────────────────┐ │
│  │ Response Dispatcher │       │ JSON Parser                  │ │
│  │ ClassifyHttp        │       │ ParseRepository              │ │
│  │ Response            │       │ ParseIssuesNodes             │ │
│  │ ClassifyGraphQL     │       │                              │ │
│  │ Errors              │       │                              │ │
│  └─────────────────────┘       └─────────────────────────────┘ │
│                                                                 │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ Utilities                                                 │   │
│  │ ParseIso8601 / FormatUnixDate / SafeListObjectName /      │   │
│  │ NullToEmpty / SafeHeader                                  │   │
│  └──────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
            │
            ▼
┌─────────────────────────────────────────┐
│ GitHub GraphQL API                       │
│ https://api.github.com/graphql           │
└─────────────────────────────────────────┘
```

### Các tầng và trách nhiệm

| Tầng | Module/Procedure | Trách nhiệm | Pure / I/O |
|------|------------------|-------------|------------|
| Public API | `GetRepoInfo`, `GetRepoIssues`, `WriteRepoIssuesTable`, `GetRepoField`, `GetRepoIssueField`, `SetGitHubToken`, `SetMaxIssues` | Điểm vào duy nhất từ ngoài; điều phối các tầng khác | I/O |
| Module state | `m_Token`, `m_MaxIssues`, `m_HttpFactory` | Lưu cấu hình trong phiên VBA | Trạng thái |
| Validation | `ValidateGithubId`, `ValidateRepoName` | Kiểm tra đầu vào (rỗng/whitespace, charset, độ dài) | Pure |
| Query Builder | `BuildRepoInfoQuery`, `BuildRepoIssuesQuery`, `BuildRequestBody`, `JsonEscape` | Xây dựng GraphQL query và JSON body | Pure |
| HTTP Transport | `SendGraphQLRequest`, `CreateHttpClient` | Gửi POST đến GitHub, đọc status/headers/body | I/O |
| Response Dispatcher | `ClassifyHttpResponse`, `ClassifyGraphQLErrors` | Phân loại mã HTTP và lỗi GraphQL thành ErrorCode | Pure |
| JSON Parser | `ParseRepository`, `ParseIssuesNodes` | Trích xuất dữ liệu từ JSON response thành UDT | Pure |
| Utilities | `ParseIso8601`, `FormatUnixDate`, `SafeListObjectName`, `NullToEmpty`, `SafeHeader` | Hàm tiện ích dùng chung | Pure |

## Components and Interfaces

### Data Types (UDTs)

```vba
' ---------------------------------------------------------------------------
' RepoInfo - Thông tin tổng quan repository
' ---------------------------------------------------------------------------
Public Type RepoInfo
    NameWithOwner   As String
    Description     As String
    PrimaryLanguage As String
    Stars           As Long
    Forks           As Long
    OpenIssues      As Long
    DefaultBranch   As String
    CreatedAt       As Date
    UpdatedAt       As Date
    Url             As String
    IsPrivate       As Boolean
    IsFork          As Boolean
    IsArchived      As Boolean
    DiskUsageKB     As Long
    ProjectsCount   As Long
    Success         As Boolean
    ErrorCode       As String
    ErrorMessage    As String
End Type

' ---------------------------------------------------------------------------
' RepoIssue - Thông tin một issue kèm project fields
' ---------------------------------------------------------------------------
Public Type RepoIssue
    Number          As Long
    Title           As String
    Url             As String
    State           As String       ' "OPEN" hoặc "CLOSED"
    Author          As String
    CreatedAt       As Date
    UpdatedAt       As Date
    ClosedAt        As Date         ' Empty nếu chưa đóng
    Labels          As String       ' Phân cách bởi ", "
    Assignees       As String       ' Phân cách bởi ", "
    Milestone       As String       ' Chuỗi rỗng nếu không có
    ProjectFields   As String       ' JSON: [{"project":"...","field":"...","value":"..."}, ...]
End Type

' ---------------------------------------------------------------------------
' RepoIssuesResult - Kết quả trả về từ GetRepoIssues
' ---------------------------------------------------------------------------
Public Type RepoIssuesResult
    Items()         As RepoIssue    ' Zero-length khi không có issue hoặc lỗi
    TotalCount      As Long
    Success         As Boolean
    ErrorCode       As String
    ErrorMessage    As String
End Type
```

### Module State

```vba
Private m_Token As String
Private m_MaxIssues As Long
Private m_HttpFactory As Object         ' Test seam for HTTP client injection
```

### Constants

```vba
Private Const DEFAULT_MAX_ISSUES As Long = 100
Private Const MIN_MAX_ISSUES As Long = 1
Private Const MAX_MAX_ISSUES As Long = 100
Private Const HTTP_TIMEOUT_MS As Long = 30000
Private Const GRAPHQL_URL As String = "https://api.github.com/graphql"
Private Const DEFAULT_USER_AGENT As String = "VBA-GitHub-GraphQL-Client"

Private Const ERR_WRITE_FAIL As Long = 1001
Private Const ERR_NULL_RANGE As Long = 1002
Private Const ERR_BAD_MAX_ISSUES As Long = 1003
```

## Interfaces

### Public API

```vba
' Cấu hình
Public Sub SetGitHubToken(ByVal token As String)
Public Sub SetMaxIssues(ByVal maxIssues As Long)
Public Sub EnsureDefaults()

' Lấy thông tin repo
Public Function GetRepoInfo(ByVal githubId As String, ByVal repoName As String) As RepoInfo
Public Function GetRepoField(ByVal githubId As String, ByVal repoName As String, ByVal fieldName As String) As Variant

' Lấy danh sách Issues
Public Function GetRepoIssues(ByVal githubId As String, ByVal repoName As String, Optional ByVal states As String = "ALL") As Variant
Public Function GetRepoIssueField(ByVal githubId As String, ByVal repoName As String, ByVal issueIndex As Long, ByVal fieldName As String) As Variant

' Ghi bảng Issues ra Excel
Public Sub WriteRepoIssuesTable(ByVal githubId As String, ByVal repoName As String, ByVal targetRange As Range, Optional ByVal columns As Variant)

' Test seam
Public Sub SetHttpFactoryForTest(ByVal factory As Object)
```

### Internal Functions

```vba
' Validation (Pure)
Private Function ValidateGithubId(ByVal s As String) As Boolean
Private Function ValidateRepoName(ByVal s As String) As Boolean

' Query Builder (Pure)
Private Function BuildRepoInfoQuery() As String
Private Function BuildRepoIssuesQuery(ByVal states As String) As String
Private Function BuildRequestBody(ByVal query As String, ByVal owner As String, ByVal name As String, Optional ByVal first As Long = -1) As String
Private Function JsonEscape(ByVal s As String) As String

' HTTP Transport (I/O)
Private Function SendGraphQLRequest(ByVal jsonBody As String, ByRef httpStatus As Long, ByRef rateLimitRemaining As String, ByRef rateLimitReset As String, ByRef responseBody As String, ByRef errCode As String, ByRef errMsg As String) As Boolean
Private Function CreateHttpClient() As Object

' Response Dispatcher (Pure)
Private Function ClassifyHttpResponse(ByVal httpStatus As Long, ByVal rateLimitRemaining As String, ByVal rateLimitReset As String, ByVal body As String, ByRef errCode As String, ByRef errMsg As String) As String
Private Function ClassifyGraphQLErrors(ByVal parsed As Object, ByVal githubId As String, ByVal repoName As String, ByRef errCode As String, ByRef errMsg As String) As Boolean

' JSON Parser (Pure)
Private Function ParseRepository(ByVal parsed As Object, ByVal githubId As String, ByVal repoName As String, ByRef result As RepoInfo) As String
Private Function ParseIssuesNodes(ByVal parsed As Object, ByRef issues() As RepoIssue, ByRef issuesCount As Long) As String

' Utilities (Pure)
Private Function ParseIso8601(ByVal s As String) As Date
Private Function FormatUnixDate(ByVal unixSeconds As Double) As String
Private Function SafeListObjectName(ByVal raw As String) As String
Private Function NullToEmpty(ByVal v As Variant) As String
Private Function SafeHeader(ByVal http As Object, ByVal headerName As String) As String
```

## Data Models

### RepoInfo UDT

| Trường | Kiểu | Mô tả |
|--------|------|-------|
| NameWithOwner | String | `owner/repo` |
| Description | String | Mô tả repo (rỗng nếu null) |
| PrimaryLanguage | String | Ngôn ngữ chính (rỗng nếu null) |
| Stars | Long | Số stargazers |
| Forks | Long | Số forks |
| OpenIssues | Long | Số issues đang mở |
| DefaultBranch | String | Tên nhánh mặc định |
| CreatedAt | Date | Ngày tạo (UTC) |
| UpdatedAt | Date | Ngày cập nhật (UTC) |
| Url | String | URL repository |
| IsPrivate | Boolean | Repo private? |
| IsFork | Boolean | Repo là fork? |
| IsArchived | Boolean | Repo đã archive? |
| DiskUsageKB | Long | Dung lượng đĩa (KB) |
| ProjectsCount | Long | Số GitHub Projects v2 |
| Success | Boolean | True nếu thành công |
| ErrorCode | String | Mã lỗi (rỗng nếu thành công) |
| ErrorMessage | String | Mô tả lỗi (rỗng nếu thành công) |

### RepoIssue UDT

| Trường | Kiểu | Mô tả |
|--------|------|-------|
| Number | Long | Số hiệu issue |
| Title | String | Tiêu đề issue |
| Url | String | URL issue trên GitHub |
| State | String | `"OPEN"` hoặc `"CLOSED"` |
| Author | String | Login của tác giả (rỗng nếu null) |
| CreatedAt | Date | Ngày tạo (UTC) |
| UpdatedAt | Date | Ngày cập nhật (UTC) |
| ClosedAt | Date | Ngày đóng (Empty nếu chưa đóng) |
| Labels | String | Danh sách label, phân cách `", "` |
| Assignees | String | Danh sách assignee login, phân cách `", "` |
| Milestone | String | Tên milestone (rỗng nếu không có) |
| ProjectFields | String | JSON array: `[{"project":"...","field":"...","value":"..."}, ...]` |

### RepoIssuesResult UDT

| Trường | Kiểu | Mô tả |
|--------|------|-------|
| Items() | RepoIssue | Mảng issues (zero-length khi rỗng/lỗi) |
| TotalCount | Long | Số lượng issue thực tế trả về |
| Success | Boolean | True nếu thành công |
| ErrorCode | String | Mã lỗi (rỗng nếu thành công) |
| ErrorMessage | String | Mô tả lỗi (rỗng nếu thành công) |

### ProjectFields JSON Format

Mỗi issue có trường `ProjectFields` chứa JSON array serialize các project field values:

```json
[
  {"project": "Sprint Board", "field": "Status", "value": "In Progress"},
  {"project": "Sprint Board", "field": "Priority", "value": "High"},
  {"project": "Roadmap", "field": "Quarter", "value": "Q1 2025"}
]
```

- `project`: Tên GitHub Project v2 mà issue thuộc
- `field`: Tên field trong project
- `value`: Giá trị field (string representation)
- Rỗng `"[]"` khi issue không thuộc project nào

## Key Design Decisions

### 1. BuildRepoIssuesQuery nhận tham số `states`

`BuildRepoIssuesQuery(states As String)` sinh ra chuỗi GraphQL query với `states` filter tương ứng:
- `"OPEN"` → `states: [OPEN]`
- `"CLOSED"` → `states: [CLOSED]`
- `"ALL"` → `states: [OPEN, CLOSED]`

Tham số `states` đã được validate và normalize (UCase) trước khi truyền vào hàm này.

### 2. ProjectFields serialization

Mỗi issue có thể thuộc nhiều projects, mỗi project có nhiều field values. Thay vì tạo UDT lồng phức tạp, thiết kế serialize toàn bộ project fields thành một chuỗi JSON đơn giản:

```json
[{"project":"Sprint Board","field":"Status","value":"In Progress"},{"project":"Sprint Board","field":"Priority","value":"High"}]
```

Chuỗi rỗng `"[]"` khi issue không thuộc project nào.

### 3. WriteRepoIssuesTable column resolution

Tham số `columns` (Optional) cho phép user chọn chính xác cột nào và thứ tự nào:
- Issue fields: 11 tên chuẩn (`"Number"`, `"Title"`, ..., `"Milestone"`)
- Project fields: format `"[ProjectTitle] FieldName"` (ví dụ `"[Sprint Board] Status"`)

Khi `columns` không truyền → xuất tất cả 11 issue fields + project fields tìm thấy (alpha sorted).

Column matching: case-insensitive cho cả issue field names và FieldName trong project field format. ProjectTitle giữ nguyên case.

### 4. GraphQL query cho Issues kèm projectItems

```graphql
query($owner: String!, $name: String!, $first: Int!) {
  repository(owner: $owner, name: $name) {
    issues(first: $first, states: [OPEN, CLOSED], orderBy: {field: CREATED_AT, direction: DESC}) {
      totalCount
      nodes {
        number
        title
        url
        state
        author { login }
        createdAt
        updatedAt
        closedAt
        labels(first: 10) { nodes { name } }
        assignees(first: 5) { nodes { login } }
        milestone { title }
        projectItems(first: 10) {
          nodes {
            project { title }
            fieldValues(first: 20) {
              nodes {
                ... on ProjectV2ItemFieldSingleSelectValue {
                  field { ... on ProjectV2SingleSelectField { name } }
                  name
                }
                ... on ProjectV2ItemFieldTextValue {
                  field { ... on ProjectV2Field { name } }
                  text
                }
                ... on ProjectV2ItemFieldNumberValue {
                  field { ... on ProjectV2Field { name } }
                  number
                }
                ... on ProjectV2ItemFieldDateValue {
                  field { ... on ProjectV2Field { name } }
                  date
                }
                ... on ProjectV2ItemFieldIterationValue {
                  field { ... on ProjectV2IterationField { name } }
                  title
                }
              }
            }
          }
        }
      }
    }
  }
}
```

### 5. Error handling nhất quán

`GetRepoIssues` reuses cùng `ClassifyHttpResponse` và `ClassifyGraphQLErrors` như `GetRepoInfo`. Thêm validation riêng cho `states` parameter (trả về `INVALID_STATES`).

### 6. SetMaxIssues validation

```vba
Public Sub SetMaxIssues(ByVal maxIssues As Long)
    If maxIssues < MIN_MAX_ISSUES Or maxIssues > MAX_MAX_ISSUES Then
        Err.Raise vbObjectError + ERR_BAD_MAX_ISSUES, "SetMaxIssues", _
                  "maxIssues must be between 1 and 100"
        Exit Sub
    End If
    m_MaxIssues = maxIssues
End Sub
```

## Data Flow

### GetRepoIssues Flow

```
GetRepoIssues(githubId, repoName, states)
  │
  ├── EnsureDefaults()  → m_MaxIssues = 100 nếu chưa set
  │
  ├── Check m_Token → MISSING_TOKEN nếu rỗng
  │
  ├── ValidateGithubId(githubId) → INVALID_GITHUB_ID nếu fail
  │
  ├── ValidateRepoName(repoName) → INVALID_REPO_NAME nếu fail
  │
  ├── Normalize states = UCase$(Trim$(states))
  │   └── Validate ∈ {"OPEN","CLOSED","ALL"} → INVALID_STATES nếu fail
  │
  ├── BuildRepoIssuesQuery(states) → query string
  │
  ├── BuildRequestBody(query, githubId, repoName, m_MaxIssues) → JSON body
  │
  ├── SendGraphQLRequest(body, ...) → httpStatus, responseBody
  │   └── Nếu fail → NETWORK_ERROR / TIMEOUT
  │
  ├── ClassifyHttpResponse(httpStatus, ...) → OK / error
  │   └── Nếu error → trả về RepoIssuesResult với ErrorCode
  │
  ├── JsonConverter.ParseJson(responseBody) → parsed Dictionary
  │   └── Nếu fail → PARSE_ERROR
  │
  ├── ClassifyGraphQLErrors(parsed, ...) → True/False
  │   └── Nếu True → trả về RepoIssuesResult với ErrorCode
  │
  └── ParseIssuesNodes(parsed, issues(), count) → "" / "PARSE_ERROR"
      └── Trả về RepoIssuesResult với Success=True, Items=issues
```

### WriteRepoIssuesTable Flow

```
WriteRepoIssuesTable(githubId, repoName, targetRange, columns)
  │
  ├── Check targetRange Is Nothing → Err.Raise ERR_NULL_RANGE
  │
  ├── result = GetRepoIssues(githubId, repoName)
  │   └── Nếu Success=False → Err.Raise ERR_WRITE_FAIL
  │
  ├── Resolve columns:
  │   ├── Nếu IsMissing(columns) → 11 issue fields + discovered project fields (alpha sorted)
  │   └── Nếu có columns → filter valid names, giữ thứ tự user chọn
  │
  ├── Ghi hàng tiêu đề tại targetRange row 1
  │
  ├── Ghi data rows (mỗi issue = 1 hàng)
  │   ├── Issue fields → đọc trực tiếp từ UDT
  │   └── Project fields → parse ProjectFields JSON, tìm match [ProjectTitle] FieldName
  │
  ├── Tạo/cập nhật ListObject
  │   ├── Nếu targetRange đã thuộc ListObject → xóa data cũ, ghi mới
  │   └── Nếu chưa → tạo ListObject tên "tblRepoIssues_{githubId}_{repoName}" (sanitized)
  │
  └── AutoFit columns
```

## Error Codes

| ErrorCode | Ngữ cảnh | Mô tả |
|-----------|----------|-------|
| `MISSING_TOKEN` | Pre-flight | Token chưa cấu hình |
| `INVALID_GITHUB_ID` | Validation | githubId rỗng/whitespace/invalid chars/quá dài |
| `INVALID_REPO_NAME` | Validation | repoName rỗng/whitespace/invalid chars/quá dài |
| `INVALID_STATES` | Validation | states không phải OPEN/CLOSED/ALL |
| `NOT_FOUND` | GraphQL | Repository không tồn tại |
| `FORBIDDEN` | HTTP/GraphQL | Không có quyền truy cập |
| `UNAUTHORIZED` | HTTP 401 | Token hết hạn/không hợp lệ |
| `RATE_LIMITED` | HTTP 403/GraphQL | Vượt quá giới hạn truy cập |
| `SERVER_ERROR` | HTTP 5xx | Lỗi phía GitHub |
| `HTTP_ERROR` | HTTP khác | Mã HTTP không xác định |
| `NETWORK_ERROR` | Transport | Lỗi kết nối/DNS/SSL |
| `TIMEOUT` | Transport | Vượt quá 30s |
| `PARSE_ERROR` | Parser | JSON không hợp lệ hoặc thiếu trường bắt buộc |
| `GRAPHQL_ERROR` | GraphQL | Lỗi GraphQL khác |

## Error Handling

### Chiến lược xử lý lỗi

1. **Không ném exception ra ngoài** (trừ `WriteRepoIssuesTable`): mọi hàm `GetRepoInfo`, `GetRepoIssues`, `GetRepoField`, `GetRepoIssueField` đều trả về giá trị có trường trạng thái hoặc `CVErr`.
2. **Fail-fast validation**: kiểm tra token → githubId → repoName → states (cho GetRepoIssues) theo thứ tự; dừng ngay khi phát hiện lỗi đầu tiên, không gọi API.
3. **Bảng mã lỗi thống nhất**: cả `GetRepoInfo` và `GetRepoIssues` dùng chung set ErrorCode (xem bảng Error Codes ở trên).
4. **WriteRepoIssuesTable**: phát sinh `Err.Raise` vì đây là Sub procedure — caller dùng `On Error` để bắt.
5. **Shared classifiers**: `ClassifyHttpResponse` và `ClassifyGraphQLErrors` được reuse giữa `GetRepoInfo` và `GetRepoIssues`.

### Error propagation flow

```
GetRepoIssues
  ├── Validation error → return RepoIssuesResult(Success=False, ErrorCode=...)
  ├── Network/Timeout  → return RepoIssuesResult(Success=False, ErrorCode="NETWORK_ERROR"/"TIMEOUT")
  ├── HTTP error       → ClassifyHttpResponse → return RepoIssuesResult(Success=False, ...)
  ├── GraphQL error    → ClassifyGraphQLErrors → return RepoIssuesResult(Success=False, ...)
  └── Parse error      → return RepoIssuesResult(Success=False, ErrorCode="PARSE_ERROR")

WriteRepoIssuesTable
  ├── targetRange=Nothing → Err.Raise vbObjectError+1002
  ├── GetRepoIssues fails → Err.Raise vbObjectError+1001
  └── Success → write data, no error
```

## Testing Strategy

### Phân tầng test

| Tầng | Chiến lược test | Công cụ |
|------|----------------|---------|
| Validation (Pure) | Property-based testing | PBT framework (VBA test module) |
| Query Builder (Pure) | Property-based testing | PBT framework |
| Response Dispatcher (Pure) | Property-based testing | PBT framework |
| JSON Parser (Pure) | Property-based + Example tests | PBT + Mock JSON fixtures |
| Utilities (Pure) | Property-based testing | PBT framework |
| HTTP Transport (I/O) | Mock-based integration tests | MockHttpClient/MockHttpFactory |
| Public API (I/O) | End-to-end mock tests | Full mock injection |
| WriteRepoIssuesTable (I/O) | Example-based integration | Excel test harness |

### Property-based test focus

Các hàm Pure là mục tiêu chính cho PBT:
- `ValidateGithubId` / `ValidateRepoName`: generate random strings, verify accept/reject
- `JsonEscape`: generate strings with special chars, verify round-trip
- `BuildRequestBody`: generate valid params, verify output is valid JSON
- `ClassifyHttpResponse`: generate all status codes, verify deterministic mapping
- `ParseIso8601`: generate valid date components, verify round-trip
- `ParseIssuesNodes`: generate mock JSON structures, verify field extraction

### Unit test focus

- `SetMaxIssues` boundary conditions: 0, 1, 100, 101
- `GetRepoIssueField` with specific fieldNames
- `WriteRepoIssuesTable` column resolution with edge cases

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system — essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Valid GitHub IDs are accepted, invalid ones are rejected

*For any* string `s`, `ValidateGithubId(s)` returns `True` if and only if `s` (after no trimming — raw input) is non-empty after `Trim$`, has length ≤ 39, and every character is in `[A-Za-z0-9-]`. Conversely, *for any* string that violates any of these conditions, `ValidateGithubId(s)` returns `False`.

**Validates: Requirements 3.1, 3.3, 17.2, 17.4**

### Property 2: Valid repo names are accepted, invalid ones are rejected

*For any* string `s`, `ValidateRepoName(s)` returns `True` if and only if `s` (after no trimming — raw input) is non-empty after `Trim$`, has length ≤ 100, and every character is in `[A-Za-z0-9._-]`. Conversely, *for any* string that violates any of these conditions, `ValidateRepoName(s)` returns `False`.

**Validates: Requirements 3.2, 3.4, 17.3, 17.5**

### Property 3: JSON escape round-trip safety

*For any* string `s`, `JsonEscape(s)` produces output that, when embedded between double quotes in a JSON document and parsed by a standard JSON parser, yields the original string `s`.

**Validates: Requirements 4.4, 15.8**

### Property 4: BuildRequestBody produces valid JSON

*For any* valid owner string, valid name string, and non-negative integer `first`, `BuildRequestBody(query, owner, name, first)` produces a string that is parseable as valid JSON and contains keys `"query"`, `"variables"` with sub-keys `"owner"`, `"name"`, and `"first"`.

**Validates: Requirements 4.3, 15.7**

### Property 5: HTTP status code classification is total and deterministic

*For any* HTTP status code in [100, 599], `ClassifyHttpResponse` maps it to exactly one of: `"OK"` (200), `"UNAUTHORIZED"` (401), `"RATE_LIMITED"` (403 + remaining=0), `"FORBIDDEN"` (403 + remaining≠0), `"SERVER_ERROR"` (500-599), or `"HTTP_ERROR"` (all others).

**Validates: Requirements 7.1, 8.1, 8.2, 8.3, 8.4, 17.12, 17.13, 17.14, 17.15, 17.16**

### Property 6: GraphQL error type classification is deterministic

*For any* parsed JSON Dictionary containing an `errors` array with at least one element, `ClassifyGraphQLErrors` maps the first error's `type` field to the correct `errCode`: `"NOT_FOUND"` → `"NOT_FOUND"`, `"FORBIDDEN"` → `"FORBIDDEN"`, `"RATE_LIMITED"` → `"RATE_LIMITED"`, any other → `"GRAPHQL_ERROR"`. When `errors` is absent/empty and `data.repository` is Nothing, it returns `"NOT_FOUND"`.

**Validates: Requirements 6.1, 6.2, 6.3, 6.4, 6.5, 17.7, 17.8, 17.9, 17.10, 17.11**

### Property 7: ParseIso8601 round-trip for valid ISO 8601 dates

*For any* valid date components (year 1900-2100, month 1-12, day 1-28, hour 0-23, minute 0-59, second 0-59) formatted as `YYYY-MM-DDTHH:MM:SSZ`, `ParseIso8601` produces a VBA Date that, when formatted back, equals the original components.

**Validates: Requirements 5.10, 5.11, 16.7, 16.8, 16.9**

### Property 8: States parameter validation and normalization

*For any* string `s`, if `UCase$(Trim$(s))` is in `{"OPEN", "CLOSED", "ALL"}` then the states validation passes; otherwise it returns `INVALID_STATES`. The comparison is case-insensitive.

**Validates: Requirements 14.8, 14.9, 14.10, 14.11**

### Property 9: BuildRepoIssuesQuery includes correct states filter

*For any* valid states value in `{"OPEN", "CLOSED", "ALL"}`, `BuildRepoIssuesQuery(states)` produces a GraphQL query string containing the corresponding states filter: `"OPEN"` → `states: [OPEN]`, `"CLOSED"` → `states: [CLOSED]`, `"ALL"` → `states: [OPEN, CLOSED]`.

**Validates: Requirements 15.2, 15.3, 15.4**

### Property 10: SetMaxIssues accepts values in [1, 100] and rejects others

*For any* Long value `n`, `SetMaxIssues(n)` succeeds (updates m_MaxIssues) if and only if `1 <= n <= 100`. For values outside this range, it raises an error and m_MaxIssues remains unchanged.

**Validates: Requirements 14.5, 14.6, 14.7**

### Property 11: ParseIssuesNodes correctly extracts issue fields from valid JSON

*For any* valid GraphQL issues response JSON (containing `data.repository.issues.nodes` array with well-formed issue objects), `ParseIssuesNodes` produces an array of `RepoIssue` where each element's `Number`, `Title`, `Url`, `State`, `Author`, `Labels`, `Assignees`, `Milestone`, and `ProjectFields` match the corresponding values in the source JSON.

**Validates: Requirements 16.1, 16.2, 16.3, 16.4, 16.5, 16.6, 16.11, 16.12, 16.13, 16.14, 16.15, 16.18, 16.19**

### Property 12: ProjectFields serialization round-trip

*For any* list of project field value triples `(projectTitle, fieldName, fieldValue)`, serializing them to JSON format `[{"project":"...","field":"...","value":"..."}, ...]` and then parsing back yields the original list of triples.

**Validates: Requirements 16.14, 16.15, 19.9**

### Property 13: GetRepoIssueField returns correct field for valid inputs

*For any* valid `fieldName` in `{"Number", "Title", "Url", "State", "Author", "CreatedAt", "UpdatedAt", "ClosedAt", "Labels", "Assignees", "Milestone", "ProjectFields"}` and valid `issueIndex` within bounds, `GetRepoIssueField` returns the value of that field from the corresponding `RepoIssue`. For invalid `fieldName`, it returns `CVErr(xlErrValue)`. For out-of-bounds `issueIndex`, it returns `CVErr(xlErrNA)`.

**Validates: Requirements 18.1, 18.2, 18.3, 18.4, 18.5, 18.6**

### Property 14: Column resolution produces correct output columns

*For any* array of column name strings and any set of issues with ProjectFields data: (a) valid issue field names (case-insensitive) map to the corresponding UDT field, (b) valid `"[ProjectTitle] FieldName"` patterns extract the matching value from ProjectFields JSON, (c) invalid column names are silently skipped, and (d) output order matches input array order.

**Validates: Requirements 19.2, 19.3, 19.4, 19.5, 19.6, 19.7, 19.9**

### Property 15: SafeListObjectName produces valid Excel ListObject names

*For any* input string, `SafeListObjectName` produces output containing only characters in `[A-Za-z0-9_]`, with length ≤ 255, and not starting with a digit.

**Validates: Requirements 19.12**
