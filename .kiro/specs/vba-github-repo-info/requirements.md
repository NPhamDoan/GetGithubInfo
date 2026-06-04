# Requirements Document

## Introduction

Tính năng này cung cấp một hàm VBA cho phép người dùng trong môi trường Microsoft Office (Excel, Word, Access) lấy thông tin về một repository trên GitHub thông qua GitHub GraphQL API. Hàm nhận hai tham số đầu vào là GitHub ID (username hoặc tên tổ chức sở hữu repo) và tên repository, sau đó gửi yêu cầu HTTP POST đến endpoint `https://api.github.com/graphql` kèm một truy vấn GraphQL được xây dựng sẵn, phân tích phản hồi JSON và trả về một đối tượng chứa các thông tin chính của repository (mô tả, ngôn ngữ chính, số sao, số fork, ngày tạo, ngày cập nhật, v.v.). Vì GitHub GraphQL API yêu cầu xác thực bắt buộc, tính năng buộc người dùng cấu hình một Personal Access Token trước khi sử dụng. Tính năng cũng xử lý các tình huống lỗi phổ biến như repository không tồn tại, lỗi mạng, vượt quá giới hạn truy cập (rate limit), lỗi xác thực, lỗi GraphQL và đầu vào không hợp lệ.

Tính năng cũng cung cấp một hàm riêng để lấy danh sách các Issues của repository kèm thông tin project fields (trạng thái, priority và các custom fields khác từ GitHub Projects v2) mà mỗi issue thuộc. Hàm này sử dụng truy vấn GraphQL với `repository.issues` kết hợp `projectItems.fieldValues` để thu thập giá trị các trường project gắn với từng issue, sau đó serialize thành chuỗi JSON. Kết quả trả về gồm mảng các đối tượng `RepoIssue` chứa đầy đủ thông tin cơ bản của issue (số hiệu, tiêu đề, URL, trạng thái, tác giả, ngày tạo/cập nhật/đóng, labels, assignees, milestone) cùng trường `ProjectFields` ở dạng JSON array. Hàm hỗ trợ lọc issues theo trạng thái (OPEN, CLOSED, hoặc ALL) qua tham số tùy chọn `states`.

## Glossary

- **GitHub_API_Client**: Mô-đun VBA chính chịu trách nhiệm gửi yêu cầu HTTP đến GitHub GraphQL API và trả kết quả về cho người gọi.
- **GetRepoInfo_Function**: Hàm VBA công khai có chữ ký `GetRepoInfo(githubId As String, repoName As String) As RepoInfo`, là điểm vào chính của tính năng.
- **RepoInfo**: Kiểu dữ liệu (Class hoặc Type trong VBA) đại diện cho thông tin repository được trả về, gồm các trường: `NameWithOwner`, `Description`, `PrimaryLanguage`, `Stars`, `Forks`, `OpenIssues`, `DefaultBranch`, `CreatedAt`, `UpdatedAt`, `Url`, `IsPrivate`, `IsFork`, `IsArchived`, `DiskUsageKB`, `ProjectsCount`, `Success`, `ErrorCode`, `ErrorMessage`.
- **JSON_Parser**: Thành phần phân tích chuỗi JSON trả về từ GitHub GraphQL API thành các giá trị VBA có thể truy cập.
- **HTTP_Client**: Đối tượng `MSXML2.XMLHTTP60` (hoặc `WinHttp.WinHttpRequest.5.1`) được sử dụng để thực hiện yêu cầu HTTP POST.
- **GitHub_GraphQL_API**: Dịch vụ web GraphQL của GitHub tại endpoint duy nhất `https://api.github.com/graphql`, sử dụng phương thức HTTP POST.
- **GraphQL_Query**: Chuỗi truy vấn GraphQL được gửi trong thân yêu cầu HTTP POST, có dạng `query($owner: String!, $name: String!) { repository(owner: $owner, name: $name) { ... } }`.
- **GraphQL_Variables**: Đối tượng JSON chứa các biến truyền vào truy vấn GraphQL, gồm `owner` (GitHub ID) và `name` (tên repository).
- **Personal_Access_Token**: Chuỗi token cá nhân do người dùng tạo trên GitHub, dùng để xác thực với GitHub_GraphQL_API. Được gắn vào header `Authorization: Bearer {token}`.
- **Rate_Limit**: Giới hạn số điểm truy vấn (point) mỗi giờ mà GitHub_GraphQL_API áp dụng cho mỗi token (mặc định 5000 điểm/giờ).
- **GitHub_ID**: Chuỗi định danh tài khoản GitHub (username hoặc tên tổ chức) sở hữu repository, được truyền vào hàm dưới dạng tham số `githubId`.
- **RepoName**: Chuỗi tên repository trên GitHub, được truyền vào hàm dưới dạng tham số `repoName`.
- **GraphQL_Errors**: Mảng `errors` trong phản hồi GraphQL, có thể xuất hiện ngay cả khi mã trạng thái HTTP là 200, mô tả các lỗi cấp GraphQL (ví dụ: `NOT_FOUND`, `FORBIDDEN`).
- **GetRepoIssues_Function**: Hàm VBA công khai có chữ ký `GetRepoIssues(githubId As String, repoName As String, Optional states As String = "ALL") As Variant`, trả về một đối tượng `RepoIssuesResult` chứa mảng các `RepoIssue` kèm thông tin project fields, là điểm vào cho tính năng lấy danh sách Issues của repository.
- **RepoIssue**: Kiểu dữ liệu (Class hoặc Type trong VBA) đại diện cho thông tin một issue của repository, gồm các trường: `Number` (Long), `Title` (String), `Url` (String), `State` (String - `"OPEN"` hoặc `"CLOSED"`), `Author` (String), `CreatedAt` (Date), `UpdatedAt` (Date), `ClosedAt` (Date hoặc Empty nếu chưa đóng), `Labels` (String - danh sách label phân cách bởi dấu phẩy), `Assignees` (String - danh sách assignee phân cách bởi dấu phẩy), `Milestone` (String - tên milestone hoặc chuỗi rỗng), `ProjectFields` (String - chuỗi JSON chứa các project field values của issue).
- **RepoIssuesResult**: Kiểu dữ liệu bọc kết quả trả về từ `GetRepoIssues_Function`, gồm các trường: `Items() As RepoIssue` (mảng các issue), `TotalCount As Long` (tổng số issue trả về), `Success As Boolean`, `ErrorCode As String`, `ErrorMessage As String`.
- **ProjectFieldValue**: Thông tin giá trị một field trong GitHub Project v2 gắn với issue, gồm `ProjectTitle` (String - tên project), `FieldName` (String - tên trường), `FieldValue` (String - giá trị trường).
- **MaxIssues**: Số lượng tối đa issue được lấy bởi `GetRepoIssues_Function` trong một lần gọi, mặc định là `100`, có thể cấu hình thông qua thủ tục `SetMaxIssues(maxIssues As Long)`.
- **WriteRepoIssuesTable_Procedure**: Thủ tục VBA công khai có chữ ký `WriteRepoIssuesTable(githubId As String, repoName As String, targetRange As Range, Optional columns As Variant)` để ghi danh sách `RepoIssue` trả về từ `GetRepoIssues_Function` ra worksheet Excel dưới dạng bảng (`ListObject`), với hàng đầu tiên là tiêu đề cột và mỗi hàng tiếp theo tương ứng một issue. Tham số `columns` (Optional, Variant chứa mảng String) cho phép user chọn chính xác các cột cần xuất và thứ tự của chúng. Tên cột hợp lệ gồm hai loại trộn chung: issue fields (`"Number"`, `"Title"`, `"Url"`, `"State"`, `"Author"`, `"CreatedAt"`, `"UpdatedAt"`, `"ClosedAt"`, `"Labels"`, `"Assignees"`, `"Milestone"`) và project fields dạng `"[ProjectTitle] FieldName"`. Nếu `columns` không truyền hoặc Empty thì xuất tất cả 11 issue fields theo thứ tự mặc định + tất cả project fields tìm thấy (sắp xếp alphabet). Không có khái niệm "cột cố định" — user toàn quyền chọn cột nào và thứ tự nào.
- **RepoIssues_Table**: Bảng Excel (`ListObject`) chứa danh sách issue trả về từ `GetRepoIssues_Function`, gồm hàng tiêu đề (các cột do user chọn qua tham số `columns` hoặc mặc định tất cả) và các hàng dữ liệu tương ứng với từng đối tượng `RepoIssue`.

## Requirements

### Requirement 1: Hàm VBA điểm vào với hai tham số

**User Story:** Là một người dùng VBA, tôi muốn gọi một hàm duy nhất với hai tham số (GitHub ID và repo name), để lấy thông tin repository GitHub mà không cần biết chi tiết về GraphQL.

#### Acceptance Criteria

1. THE GetRepoInfo_Function SHALL chấp nhận đúng hai tham số kiểu `String`: `githubId` và `repoName`, theo thứ tự đó.
2. THE GetRepoInfo_Function SHALL trả về một đối tượng kiểu `RepoInfo`.
3. WHEN GetRepoInfo_Function được gọi với cả hai tham số hợp lệ và Personal_Access_Token đã được cấu hình, THE GitHub_API_Client SHALL gửi yêu cầu HTTP POST đến URL cố định `https://api.github.com/graphql`.
4. THE GitHub_API_Client SHALL gắn header `Content-Type: application/json` vào mọi yêu cầu HTTP gửi đến GitHub_GraphQL_API.
5. THE GitHub_API_Client SHALL gắn header `Accept: application/json` vào mọi yêu cầu HTTP gửi đến GitHub_GraphQL_API.
6. THE GitHub_API_Client SHALL gắn header `User-Agent` với giá trị không rỗng (mặc định `VBA-GitHub-GraphQL-Client`) vào mọi yêu cầu HTTP gửi đến GitHub_GraphQL_API.

### Requirement 2: Yêu cầu xác thực bằng Personal Access Token

**User Story:** Là một người dùng VBA, tôi muốn cấu hình Personal Access Token một lần và được nhắc rõ ràng nếu chưa cấu hình, để tuân thủ yêu cầu xác thực bắt buộc của GitHub GraphQL API.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL cung cấp thủ tục công khai `SetGitHubToken(token As String)` để người dùng cấu hình Personal_Access_Token trong phiên VBA.
2. WHEN một yêu cầu HTTP được gửi đến GitHub_GraphQL_API, THE GitHub_API_Client SHALL gắn header `Authorization: Bearer {token}` với giá trị token đã cấu hình.
3. IF Personal_Access_Token chưa được cấu hình hoặc là chuỗi rỗng tại thời điểm gọi GetRepoInfo_Function, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "MISSING_TOKEN"` và `ErrorMessage` mô tả yêu cầu cấu hình token trước khi sử dụng.
4. WHEN GetRepoInfo_Function trả về `ErrorCode = "MISSING_TOKEN"`, THE GitHub_API_Client SHALL không gửi yêu cầu HTTP đến GitHub_GraphQL_API.

### Requirement 3: Xác thực đầu vào

**User Story:** Là một người dùng VBA, tôi muốn hàm phát hiện đầu vào không hợp lệ ngay lập tức, để tránh gọi API không cần thiết và nhận được thông báo lỗi rõ ràng.

#### Acceptance Criteria

1. IF tham số `githubId` là chuỗi rỗng hoặc chỉ chứa ký tự khoảng trắng, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "INVALID_GITHUB_ID"` và `ErrorMessage` mô tả tham số `githubId` không hợp lệ.
2. IF tham số `repoName` là chuỗi rỗng hoặc chỉ chứa ký tự khoảng trắng, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "INVALID_REPO_NAME"` và `ErrorMessage` mô tả tham số `repoName` không hợp lệ.
3. IF tham số `githubId` chứa ký tự ngoài tập `[A-Za-z0-9-]` hoặc dài hơn 39 ký tự, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "INVALID_GITHUB_ID"`.
4. IF tham số `repoName` chứa ký tự ngoài tập `[A-Za-z0-9._-]` hoặc dài hơn 100 ký tự, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "INVALID_REPO_NAME"`.
5. WHEN GetRepoInfo_Function phát hiện đầu vào không hợp lệ, THE GitHub_API_Client SHALL không gửi yêu cầu HTTP đến GitHub_GraphQL_API.

### Requirement 4: Xây dựng truy vấn GraphQL và thân yêu cầu

**User Story:** Là một người dùng VBA, tôi muốn tham số được truyền vào GraphQL một cách an toàn qua biến (variables), để tránh lỗi cú pháp và nguy cơ injection.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL sử dụng truy vấn GraphQL có tham số hóa với hai biến `$owner: String!` và `$name: String!` thay vì nội suy chuỗi trực tiếp vào thân truy vấn.
2. THE GraphQL_Query SHALL yêu cầu các trường tối thiểu sau từ object `repository`: `nameWithOwner`, `description`, `url`, `stargazerCount`, `forkCount`, `isPrivate`, `isFork`, `isArchived`, `diskUsage`, `createdAt`, `updatedAt`, `defaultBranchRef { name }`, `primaryLanguage { name }`, `issues(states: OPEN) { totalCount }`, `projectsV2 { totalCount }`.
3. WHEN GitHub_API_Client xây dựng thân yêu cầu HTTP POST, THE GitHub_API_Client SHALL nối GraphQL_Query và GraphQL_Variables thành một đối tượng JSON có cấu trúc `{ "query": "...", "variables": { "owner": githubId, "name": repoName } }`.
4. THE GitHub_API_Client SHALL áp dụng escape JSON cho giá trị `githubId` và `repoName` khi xây dựng thân yêu cầu để các ký tự đặc biệt không phá vỡ cấu trúc JSON.

### Requirement 5: Phân tích JSON phản hồi thành công

**User Story:** Là một người dùng VBA, tôi muốn nhận được các trường thông tin repository ở dạng có cấu trúc, để có thể sử dụng trực tiếp trong macro hoặc bảng tính.

#### Acceptance Criteria

1. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và phản hồi không chứa GraphQL_Errors và trường `data.repository` khác `null`, THE JSON_Parser SHALL phân tích nội dung phản hồi thành các trường của RepoInfo.
2. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.nameWithOwner` vào trường `RepoInfo.NameWithOwner`.
3. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.description` vào trường `RepoInfo.Description`, và gán chuỗi rỗng nếu giá trị JSON là `null`.
4. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.primaryLanguage.name` vào trường `RepoInfo.PrimaryLanguage`, và gán chuỗi rỗng nếu `data.repository.primaryLanguage` là `null`.
5. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.stargazerCount` vào trường `RepoInfo.Stars` dưới dạng số nguyên `Long`.
6. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.forkCount` vào trường `RepoInfo.Forks` dưới dạng số nguyên `Long`.
7. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.issues.totalCount` vào trường `RepoInfo.OpenIssues` dưới dạng số nguyên `Long`.
8. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.defaultBranchRef.name` vào trường `RepoInfo.DefaultBranch`, và gán chuỗi rỗng nếu `data.repository.defaultBranchRef` là `null`.
9. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.url` vào trường `RepoInfo.Url`.
10. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL chuyển giá trị `data.repository.createdAt` (định dạng ISO 8601) thành kiểu `Date` của VBA và gán vào trường `RepoInfo.CreatedAt`.
11. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL chuyển giá trị `data.repository.updatedAt` (định dạng ISO 8601) thành kiểu `Date` của VBA và gán vào trường `RepoInfo.UpdatedAt`.
12. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị boolean `data.repository.isPrivate` vào trường `RepoInfo.IsPrivate`.
13. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị boolean `data.repository.isFork` vào trường `RepoInfo.IsFork`.
14. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị boolean `data.repository.isArchived` vào trường `RepoInfo.IsArchived`.
15. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.diskUsage` vào trường `RepoInfo.DiskUsageKB` dưới dạng số nguyên `Long`.
16. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL gán giá trị `data.repository.projectsV2.totalCount` vào trường `RepoInfo.ProjectsCount` dưới dạng số nguyên `Long`, và gán `0` nếu `data.repository.projectsV2` là `null`.
17. WHEN phân tích JSON thành công, THE GetRepoInfo_Function SHALL đặt `RepoInfo.Success = True`, `RepoInfo.ErrorCode = ""` và `RepoInfo.ErrorMessage = ""`.

### Requirement 6: Xử lý lỗi GraphQL cấp ứng dụng

**User Story:** Là một người dùng VBA, tôi muốn các lỗi do GraphQL trả về được phân loại đúng, để có thể chẩn đoán nguyên nhân kể cả khi mã trạng thái HTTP là 200.

#### Acceptance Criteria

1. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và `data.repository` là `null` mà không có mục nào trong mảng `errors`, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "NOT_FOUND"` và `ErrorMessage` chứa chuỗi `githubId/repoName` đã yêu cầu.
2. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục có `type = "NOT_FOUND"`, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "NOT_FOUND"` và `ErrorMessage` chứa chuỗi `githubId/repoName` đã yêu cầu.
3. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục có `type = "FORBIDDEN"`, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "FORBIDDEN"` và `ErrorMessage` chứa giá trị `message` của mục lỗi đầu tiên.
4. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục có `type = "RATE_LIMITED"`, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "RATE_LIMITED"` và `ErrorMessage` chứa giá trị `message` của mục lỗi đầu tiên.
5. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục mà không khớp với các trường hợp 6.2 đến 6.4, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "GRAPHQL_ERROR"` và `ErrorMessage` chứa giá trị `message` của mục lỗi đầu tiên.

### Requirement 7: Xử lý giới hạn tốc độ truy cập ở cấp HTTP

**User Story:** Là một người dùng VBA, tôi muốn biết khi yêu cầu bị chặn do vượt quá giới hạn tốc độ truy cập của GitHub, để có thể chờ và thử lại sau.

#### Acceptance Criteria

1. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 403 và header phản hồi `X-RateLimit-Remaining` có giá trị `0`, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "RATE_LIMITED"`.
2. WHEN GetRepoInfo_Function trả về `ErrorCode = "RATE_LIMITED"` do mã trạng thái HTTP 403, THE GetRepoInfo_Function SHALL điền vào `ErrorMessage` thời điểm đặt lại giới hạn lấy từ header `X-RateLimit-Reset` (chuyển từ Unix timestamp sang định dạng ngày giờ địa phương).

### Requirement 8: Xử lý lỗi xác thực và lỗi máy chủ ở cấp HTTP

**User Story:** Là một người dùng VBA, tôi muốn các lỗi HTTP khác được phân biệt rõ ràng, để có thể chẩn đoán nguyên nhân.

#### Acceptance Criteria

1. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 401, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "UNAUTHORIZED"` và `ErrorMessage` mô tả token không hợp lệ hoặc đã hết hạn.
2. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 403 và header `X-RateLimit-Remaining` khác `0` hoặc không có, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "FORBIDDEN"`.
3. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP trong khoảng 500 đến 599, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "SERVER_ERROR"`.
4. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP khác các trường hợp đã định nghĩa ở các requirement 5, 6, 7 và 8.1 đến 8.3, THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False`, `ErrorCode = "HTTP_ERROR"` và `ErrorMessage` chứa mã trạng thái HTTP nhận được.

### Requirement 9: Xử lý lỗi mạng

**User Story:** Là một người dùng VBA, tôi muốn hàm không làm treo macro khi có sự cố mạng, để Excel/Word vẫn hoạt động bình thường.

#### Acceptance Criteria

1. IF HTTP_Client phát sinh lỗi (ví dụ không có kết nối Internet, DNS không phân giải, lỗi SSL), THEN THE GetRepoInfo_Function SHALL bắt lỗi và trả về một `RepoInfo` với `Success = False`, `ErrorCode = "NETWORK_ERROR"` và `ErrorMessage` chứa mô tả lỗi từ `Err.Description`.
2. THE GitHub_API_Client SHALL áp dụng thời gian chờ tối đa 30 giây cho mỗi yêu cầu HTTP đến GitHub_GraphQL_API.
3. IF yêu cầu HTTP vượt quá thời gian chờ 30 giây, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "TIMEOUT"`.

### Requirement 10: Xử lý lỗi phân tích JSON

**User Story:** Là một người dùng VBA, tôi muốn hàm không bị crash khi GitHub trả về phản hồi không đúng định dạng, để macro vẫn kết thúc an toàn.

#### Acceptance Criteria

1. IF nội dung phản hồi từ GitHub_GraphQL_API không phải là JSON hợp lệ, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "PARSE_ERROR"`.
2. IF JSON phản hồi thành công không chứa cả hai khóa `data` và `repository` ở vị trí mong đợi, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "PARSE_ERROR"`.
3. IF JSON phản hồi thành công thiếu một trường bắt buộc trong số `nameWithOwner`, `createdAt`, `updatedAt`, THEN THE GetRepoInfo_Function SHALL trả về một `RepoInfo` với `Success = False` và `ErrorCode = "PARSE_ERROR"`.

### Requirement 11: Tương thích môi trường VBA

**User Story:** Là một người dùng VBA, tôi muốn sử dụng tính năng này từ Excel, Word hoặc Access, để có thể tích hợp vào nhiều loại tài liệu Office.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL được triển khai dưới dạng mô-đun chuẩn (`.bas`) có thể nhập vào Excel, Word và Access VBA mà không cần thay đổi mã nguồn.
2. THE GitHub_API_Client SHALL không phụ thuộc vào tham chiếu thư viện đòi hỏi cài đặt thủ công ngoài các thư viện có sẵn trên Windows (ví dụ `Microsoft XML, v6.0`, `Microsoft Scripting Runtime`).
3. WHERE thư viện `Microsoft Scripting Runtime` chưa được tham chiếu, THE GitHub_API_Client SHALL sử dụng `CreateObject` thay cho early binding để tránh lỗi biên dịch.

### Requirement 12: Sử dụng từ Excel worksheet

**User Story:** Là một người dùng Excel, tôi muốn gọi hàm trực tiếp từ ô bảng tính, để hiển thị thông tin repository ngay trong sheet.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL cung cấp một hàm phụ `GetRepoField(githubId As String, repoName As String, fieldName As String) As Variant` có thể được gọi như công thức trong ô Excel.
2. WHEN GetRepoField được gọi với `fieldName` là một trong các giá trị `"NameWithOwner"`, `"Description"`, `"PrimaryLanguage"`, `"Stars"`, `"Forks"`, `"OpenIssues"`, `"DefaultBranch"`, `"Url"`, `"CreatedAt"`, `"UpdatedAt"`, `"IsPrivate"`, `"IsFork"`, `"IsArchived"`, `"DiskUsageKB"`, `"ProjectsCount"`, THE GetRepoField SHALL trả về giá trị tương ứng của trường đó từ `RepoInfo`.
3. IF `fieldName` không nằm trong danh sách trường được hỗ trợ, THEN THE GetRepoField SHALL trả về giá trị lỗi `CVErr(xlErrValue)`.
4. IF `RepoInfo.Success = False`, THEN THE GetRepoField SHALL trả về giá trị lỗi `CVErr(xlErrNA)`.

### Requirement 13: Tài liệu sử dụng

**User Story:** Là một người dùng VBA mới, tôi muốn có hướng dẫn sử dụng đi kèm, để có thể áp dụng tính năng nhanh chóng.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL đi kèm tệp `README.md` mô tả cách nhập mô-đun vào Excel/Word/Access, các tham chiếu thư viện cần thiết, cách tạo Personal Access Token trên GitHub và cách cấu hình token qua `SetGitHubToken`.
2. THE README.md SHALL bao gồm ít nhất một ví dụ gọi `GetRepoInfo` từ `Sub` VBA và một ví dụ gọi `GetRepoField` từ ô Excel.
3. THE README.md SHALL liệt kê đầy đủ các giá trị có thể có của `ErrorCode` cùng ý nghĩa của từng giá trị.
4. THE README.md SHALL mô tả tính năng `GetRepoIssues_Function`, `GetRepoIssueField` và `WriteRepoIssuesTable`, kèm ít nhất một ví dụ gọi `GetRepoIssues` từ `Sub` VBA và một ví dụ gọi `GetRepoIssueField` từ ô Excel.
5. THE README.md SHALL nêu rõ rằng Personal_Access_Token cần có scope `read:project` (hoặc `project` cho quyền ghi) để sử dụng `GetRepoIssues_Function` với thông tin project fields và trường `RepoInfo.ProjectsCount`.
6. THE README.md SHALL giải thích cấu trúc dữ liệu `ProjectFields` (JSON array dạng `[{"project":"...","field":"...","value":"..."}, ...]`) và cách parse chuỗi JSON này trong VBA nếu cần truy cập từng field value riêng lẻ.
7. THE README.md SHALL mô tả cách sử dụng `SetMaxIssues` để cấu hình số lượng issue tối đa (mặc định 100, phạm vi [1, 100]).

### Requirement 14: Hàm VBA lấy danh sách Issues kèm project fields

**User Story:** Là một người dùng VBA, tôi muốn gọi một hàm với GitHub ID và repo name để lấy danh sách issues kèm thông tin project fields, để biết trạng thái/priority của từng issue trong project.

#### Acceptance Criteria

1. THE GetRepoIssues_Function SHALL chấp nhận ba tham số: `githubId As String`, `repoName As String` và `Optional states As String = "ALL"`, theo thứ tự đó.
2. THE GetRepoIssues_Function SHALL trả về một giá trị kiểu `Variant` chứa đối tượng `RepoIssuesResult` với mảng các `RepoIssue` được sắp xếp theo thứ tự xuất hiện trong phản hồi GraphQL (mới nhất trước).
3. WHEN GetRepoIssues_Function được gọi với cả hai tham số bắt buộc hợp lệ và Personal_Access_Token đã được cấu hình, THE GitHub_API_Client SHALL gửi yêu cầu HTTP POST đến URL cố định `https://api.github.com/graphql`.
4. THE GitHub_API_Client SHALL gắn các header `Content-Type: application/json`, `Accept: application/json`, `User-Agent` (giá trị không rỗng) và `Authorization: Bearer {token}` vào mọi yêu cầu HTTP gửi đến GitHub_GraphQL_API thông qua GetRepoIssues_Function, tương tự như đối với GetRepoInfo_Function.
5. THE GitHub_API_Client SHALL cung cấp thủ tục công khai `SetMaxIssues(maxIssues As Long)` để người dùng cấu hình số lượng issue tối đa được lấy bởi GetRepoIssues_Function trong một lần gọi.
6. WHEN người dùng chưa gọi `SetMaxIssues`, THE GetRepoIssues_Function SHALL sử dụng giá trị mặc định MaxIssues bằng `100`.
7. IF tham số `maxIssues` truyền vào `SetMaxIssues` nhỏ hơn `1` hoặc lớn hơn `100`, THEN THE GitHub_API_Client SHALL giữ nguyên giá trị MaxIssues hiện tại và phát sinh giá trị Err với mô tả phạm vi cho phép `[1, 100]`.
8. WHEN tham số `states` có giá trị `"OPEN"` (case-insensitive), THE GetRepoIssues_Function SHALL chỉ lấy các issue có trạng thái OPEN.
9. WHEN tham số `states` có giá trị `"CLOSED"` (case-insensitive), THE GetRepoIssues_Function SHALL chỉ lấy các issue có trạng thái CLOSED.
10. WHEN tham số `states` có giá trị `"ALL"` (case-insensitive) hoặc không được truyền, THE GetRepoIssues_Function SHALL lấy tất cả issues không phân biệt trạng thái.
11. IF tham số `states` có giá trị khác `"OPEN"`, `"CLOSED"` hoặc `"ALL"` (case-insensitive), THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "INVALID_STATES"` và `ErrorMessage` mô tả giá trị `states` không hợp lệ.

### Requirement 15: Xây dựng truy vấn GraphQL cho Issues kèm projectItems

**User Story:** Là một người dùng VBA, tôi muốn hàm GetRepoIssues sử dụng truy vấn GraphQL đúng chuẩn cho `repository.issues` kết hợp `projectItems.fieldValues`, để lấy đầy đủ thông tin issue cùng giá trị các trường project.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL sử dụng truy vấn GraphQL có tham số hóa với ba biến `$owner: String!`, `$name: String!` và `$first: Int!` thay vì nội suy chuỗi trực tiếp vào thân truy vấn.
2. WHEN tham số `states` là `"OPEN"`, THE GraphQL_Query cho GetRepoIssues_Function SHALL truy vấn `repository.issues(first: $first, states: [OPEN], orderBy: {field: CREATED_AT, direction: DESC})`.
3. WHEN tham số `states` là `"CLOSED"`, THE GraphQL_Query cho GetRepoIssues_Function SHALL truy vấn `repository.issues(first: $first, states: [CLOSED], orderBy: {field: CREATED_AT, direction: DESC})`.
4. WHEN tham số `states` là `"ALL"`, THE GraphQL_Query cho GetRepoIssues_Function SHALL truy vấn `repository.issues(first: $first, states: [OPEN, CLOSED], orderBy: {field: CREATED_AT, direction: DESC})`.
5. THE GraphQL_Query cho GetRepoIssues_Function SHALL yêu cầu `totalCount` và `nodes` chứa các trường: `number`, `title`, `url`, `state`, `author { login }`, `createdAt`, `updatedAt`, `closedAt`, `labels(first: 10) { nodes { name } }`, `assignees(first: 5) { nodes { login } }`, `milestone { title }`.
6. THE GraphQL_Query cho GetRepoIssues_Function SHALL bao gồm trường `projectItems(first: 10) { nodes { project { title } fieldValues(first: 20) { nodes { ... on ProjectV2ItemFieldSingleSelectValue { field { ... on ProjectV2SingleSelectField { name } } name optionId } ... on ProjectV2ItemFieldTextValue { field { ... on ProjectV2Field { name } } text } ... on ProjectV2ItemFieldNumberValue { field { ... on ProjectV2Field { name } } number } ... on ProjectV2ItemFieldDateValue { field { ... on ProjectV2Field { name } } date } ... on ProjectV2ItemFieldIterationValue { field { ... on ProjectV2IterationField { name } } title } } } } }` trong mỗi node issue.
7. WHEN GitHub_API_Client xây dựng thân yêu cầu HTTP POST cho GetRepoIssues_Function, THE GitHub_API_Client SHALL nối GraphQL_Query và GraphQL_Variables thành một đối tượng JSON có cấu trúc `{ "query": "...", "variables": { "owner": githubId, "name": repoName, "first": MaxIssues } }`.
8. THE GitHub_API_Client SHALL áp dụng escape JSON cho giá trị `githubId` và `repoName` khi xây dựng thân yêu cầu để các ký tự đặc biệt không phá vỡ cấu trúc JSON.

### Requirement 16: Phân tích JSON phản hồi Issues

**User Story:** Là một người dùng VBA, tôi muốn nhận được danh sách issues ở dạng có cấu trúc với thông tin project fields đã được serialize, để có thể duyệt và sử dụng trực tiếp trong macro hoặc bảng tính.

#### Acceptance Criteria

1. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200, phản hồi không chứa GraphQL_Errors và `data.repository` khác `null`, THE JSON_Parser SHALL phân tích nội dung phản hồi thành mảng các đối tượng `RepoIssue` từ `data.repository.issues.nodes`.
2. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL gán giá trị `nodes[i].number` vào trường `RepoIssue.Number` của phần tử thứ `i` dưới dạng số nguyên `Long`.
3. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL gán giá trị `nodes[i].title` vào trường `RepoIssue.Title` của phần tử thứ `i`.
4. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL gán giá trị `nodes[i].url` vào trường `RepoIssue.Url` của phần tử thứ `i`.
5. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL gán giá trị `nodes[i].state` (chuỗi `"OPEN"` hoặc `"CLOSED"`) vào trường `RepoIssue.State` của phần tử thứ `i`.
6. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL gán giá trị `nodes[i].author.login` vào trường `RepoIssue.Author` của phần tử thứ `i`, và gán chuỗi rỗng nếu `nodes[i].author` là `null`.
7. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL chuyển giá trị `nodes[i].createdAt` (định dạng ISO 8601) thành kiểu `Date` của VBA và gán vào trường `RepoIssue.CreatedAt` của phần tử thứ `i`.
8. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL chuyển giá trị `nodes[i].updatedAt` (định dạng ISO 8601) thành kiểu `Date` của VBA và gán vào trường `RepoIssue.UpdatedAt` của phần tử thứ `i`.
9. WHEN phân tích JSON thành công và `nodes[i].closedAt` khác `null`, THE GetRepoIssues_Function SHALL chuyển giá trị `nodes[i].closedAt` (định dạng ISO 8601) thành kiểu `Date` của VBA và gán vào trường `RepoIssue.ClosedAt` của phần tử thứ `i`.
10. WHEN phân tích JSON thành công và `nodes[i].closedAt` là `null`, THE GetRepoIssues_Function SHALL gán giá trị `Empty` vào trường `RepoIssue.ClosedAt` của phần tử thứ `i`.
11. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL nối tên các label trong `nodes[i].labels.nodes` bằng dấu phẩy và khoảng trắng (`", "`) và gán vào trường `RepoIssue.Labels` của phần tử thứ `i`, gán chuỗi rỗng nếu không có label nào.
12. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL nối login của các assignee trong `nodes[i].assignees.nodes` bằng dấu phẩy và khoảng trắng (`", "`) và gán vào trường `RepoIssue.Assignees` của phần tử thứ `i`, gán chuỗi rỗng nếu không có assignee nào.
13. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL gán giá trị `nodes[i].milestone.title` vào trường `RepoIssue.Milestone` của phần tử thứ `i`, và gán chuỗi rỗng nếu `nodes[i].milestone` là `null`.
14. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL duyệt `nodes[i].projectItems.nodes` và cho mỗi project item, trích xuất `project.title` cùng các `fieldValues.nodes` để tạo danh sách các `ProjectFieldValue`, sau đó serialize thành chuỗi JSON dạng `[{"project":"ProjectTitle","field":"FieldName","value":"FieldValue"}, ...]` và gán vào trường `RepoIssue.ProjectFields` của phần tử thứ `i`.
15. WHEN phân tích JSON thành công và issue không thuộc project nào (mảng `projectItems.nodes` rỗng), THE GetRepoIssues_Function SHALL gán chuỗi `"[]"` vào trường `RepoIssue.ProjectFields` của phần tử thứ `i`.
16. WHEN phân tích JSON thành công và `data.repository.issues.totalCount` bằng `0`, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `TotalCount = 0`, `Success = True`, `ErrorCode = ""` và `ErrorMessage = ""`.
17. WHEN phân tích JSON thành công và `data.repository.issues.nodes` là mảng rỗng, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `TotalCount = 0`, `Success = True`, `ErrorCode = ""` và `ErrorMessage = ""`.
18. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL gán `RepoIssuesResult.TotalCount` bằng số lượng phần tử thực tế trong mảng `Items` đã phân tích được.
19. WHEN phân tích JSON thành công, THE GetRepoIssues_Function SHALL đặt `RepoIssuesResult.Success = True`, `RepoIssuesResult.ErrorCode = ""` và `RepoIssuesResult.ErrorMessage = ""`.

### Requirement 17: Xử lý lỗi cho GetRepoIssues

**User Story:** Là một người dùng VBA, tôi muốn các lỗi xác thực đầu vào, GraphQL, HTTP, mạng, timeout và phân tích JSON khi gọi GetRepoIssues được phân loại nhất quán với GetRepoInfo, để có thể xử lý lỗi theo cùng một cách.

#### Acceptance Criteria

1. IF Personal_Access_Token chưa được cấu hình hoặc là chuỗi rỗng tại thời điểm gọi GetRepoIssues_Function, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "MISSING_TOKEN"` và `ErrorMessage` mô tả yêu cầu cấu hình token trước khi sử dụng.
2. IF tham số `githubId` là chuỗi rỗng hoặc chỉ chứa ký tự khoảng trắng, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "INVALID_GITHUB_ID"` và `ErrorMessage` mô tả tham số `githubId` không hợp lệ.
3. IF tham số `repoName` là chuỗi rỗng hoặc chỉ chứa ký tự khoảng trắng, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "INVALID_REPO_NAME"` và `ErrorMessage` mô tả tham số `repoName` không hợp lệ.
4. IF tham số `githubId` chứa ký tự ngoài tập `[A-Za-z0-9-]` hoặc dài hơn 39 ký tự, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "INVALID_GITHUB_ID"`.
5. IF tham số `repoName` chứa ký tự ngoài tập `[A-Za-z0-9._-]` hoặc dài hơn 100 ký tự, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "INVALID_REPO_NAME"`.
6. WHEN GetRepoIssues_Function phát hiện đầu vào không hợp lệ hoặc thiếu Personal_Access_Token, THE GitHub_API_Client SHALL không gửi yêu cầu HTTP đến GitHub_GraphQL_API.
7. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và `data.repository` là `null` mà không có mục nào trong mảng `errors`, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "NOT_FOUND"` và `ErrorMessage` chứa chuỗi `githubId/repoName` đã yêu cầu.
8. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục có `type = "NOT_FOUND"`, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "NOT_FOUND"` và `ErrorMessage` chứa chuỗi `githubId/repoName` đã yêu cầu.
9. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục có `type = "FORBIDDEN"`, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "FORBIDDEN"` và `ErrorMessage` chứa giá trị `message` của mục lỗi đầu tiên.
10. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục có `type = "RATE_LIMITED"`, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "RATE_LIMITED"` và `ErrorMessage` chứa giá trị `message` của mục lỗi đầu tiên.
11. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 200 và mảng `errors` chứa ít nhất một mục mà không khớp với các trường hợp 17.8 đến 17.10, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "GRAPHQL_ERROR"` và `ErrorMessage` chứa giá trị `message` của mục lỗi đầu tiên.
12. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 401, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "UNAUTHORIZED"` và `ErrorMessage` mô tả token không hợp lệ hoặc đã hết hạn.
13. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 403 và header phản hồi `X-RateLimit-Remaining` có giá trị `0`, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "RATE_LIMITED"`, với `ErrorMessage` chứa thời điểm đặt lại giới hạn lấy từ header `X-RateLimit-Reset` (chuyển từ Unix timestamp sang định dạng ngày giờ địa phương).
14. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP 403 và header `X-RateLimit-Remaining` khác `0` hoặc không có, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "FORBIDDEN"`.
15. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP trong khoảng 500 đến 599, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "SERVER_ERROR"`.
16. WHEN GitHub_GraphQL_API trả về mã trạng thái HTTP khác các trường hợp đã định nghĩa ở các requirement 16 và 17.7 đến 17.15, THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "HTTP_ERROR"` và `ErrorMessage` chứa mã trạng thái HTTP nhận được.
17. IF HTTP_Client phát sinh lỗi (ví dụ không có kết nối Internet, DNS không phân giải, lỗi SSL), THEN THE GetRepoIssues_Function SHALL bắt lỗi và trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False`, `ErrorCode = "NETWORK_ERROR"` và `ErrorMessage` chứa mô tả lỗi từ `Err.Description`.
18. THE GitHub_API_Client SHALL áp dụng thời gian chờ tối đa 30 giây cho yêu cầu HTTP của GetRepoIssues_Function đến GitHub_GraphQL_API.
19. IF yêu cầu HTTP của GetRepoIssues_Function vượt quá thời gian chờ 30 giây, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "TIMEOUT"`.
20. IF nội dung phản hồi từ GitHub_GraphQL_API không phải là JSON hợp lệ, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "PARSE_ERROR"`.
21. IF JSON phản hồi thành công không chứa cả ba khóa `data`, `repository` và `issues` ở vị trí mong đợi, THEN THE GetRepoIssues_Function SHALL trả về `RepoIssuesResult` với mảng `Items` rỗng, `Success = False` và `ErrorCode = "PARSE_ERROR"`.

### Requirement 18: Sử dụng GetRepoIssues từ Excel worksheet

**User Story:** Là một người dùng Excel, tôi muốn truy cập từng trường thông tin của một issue trực tiếp từ ô bảng tính, để hiển thị danh sách issue trong sheet.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL cung cấp một hàm phụ `GetRepoIssueField(githubId As String, repoName As String, issueIndex As Long, fieldName As String) As Variant` có thể được gọi như công thức trong ô Excel.
2. THE GetRepoIssueField SHALL sử dụng chỉ số `issueIndex` bắt đầu từ `1` để tham chiếu phần tử đầu tiên của mảng `RepoIssue` trả về bởi GetRepoIssues_Function.
3. WHEN GetRepoIssueField được gọi với `fieldName` là một trong các giá trị `"Number"`, `"Title"`, `"Url"`, `"State"`, `"Author"`, `"CreatedAt"`, `"UpdatedAt"`, `"ClosedAt"`, `"Labels"`, `"Assignees"`, `"Milestone"`, `"ProjectFields"` và `issueIndex` nằm trong phạm vi hợp lệ, THE GetRepoIssueField SHALL trả về giá trị tương ứng của trường đó từ phần tử `RepoIssue` thứ `issueIndex`.
4. IF `fieldName` không nằm trong danh sách trường được hỗ trợ, THEN THE GetRepoIssueField SHALL trả về giá trị lỗi `CVErr(xlErrValue)`.
5. IF kết quả của GetRepoIssues_Function có `Success = False`, THEN THE GetRepoIssueField SHALL trả về giá trị lỗi `CVErr(xlErrNA)`.
6. IF `issueIndex` nhỏ hơn `1` hoặc lớn hơn số lượng phần tử trong mảng `RepoIssue` trả về, THEN THE GetRepoIssueField SHALL trả về giá trị lỗi `CVErr(xlErrNA)`.

### Requirement 19: Xuất danh sách Issues ra bảng trên Excel worksheet với cột tùy chọn

**User Story:** Là một người dùng Excel, tôi muốn ghi danh sách issue trả về từ `GetRepoIssues` ra một bảng trên worksheet và toàn quyền chọn cột nào cần xuất cùng thứ tự của chúng (bao gồm cả issue fields và project fields), để xem trực quan thông tin issues trong bảng tính mà không bị ràng buộc bởi cột cố định.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL cung cấp thủ tục công khai `WriteRepoIssuesTable(githubId As String, repoName As String, targetRange As Range, Optional columns As Variant)` để ghi kết quả của GetRepoIssues_Function ra một vùng ô trên worksheet, bắt đầu tại ô trên-trái của `targetRange`.
2. THE WriteRepoIssuesTable_Procedure SHALL chấp nhận tham số `columns` kiểu `Optional Variant` chứa mảng String — danh sách tên cột muốn xuất; tên cột hợp lệ gồm hai loại trộn chung trong cùng một mảng: issue fields (`"Number"`, `"Title"`, `"Url"`, `"State"`, `"Author"`, `"CreatedAt"`, `"UpdatedAt"`, `"ClosedAt"`, `"Labels"`, `"Assignees"`, `"Milestone"`) và project fields dạng `"[ProjectTitle] FieldName"` (ví dụ `"[My Board] Status"`, `"[Sprint] Priority"`).
3. WHEN tham số `columns` không được truyền hoặc có giá trị `Empty` hoặc `IsMissing(columns) = True`, THE WriteRepoIssuesTable_Procedure SHALL xuất TẤT CẢ các cột: 11 issue fields theo thứ tự mặc định (`"Number"`, `"Title"`, `"Url"`, `"State"`, `"Author"`, `"CreatedAt"`, `"UpdatedAt"`, `"ClosedAt"`, `"Labels"`, `"Assignees"`, `"Milestone"`) tiếp theo là tất cả project fields tìm thấy (quét tất cả issues, thu thập danh sách unique `"[ProjectTitle] FieldName"` và sắp xếp alphabet).
4. WHEN tham số `columns` được truyền dưới dạng mảng String, THE WriteRepoIssuesTable_Procedure SHALL chỉ xuất các cột đã liệt kê trong mảng, theo đúng thứ tự xuất hiện trong mảng đó — không có khái niệm "cột cố định".
5. WHEN WriteRepoIssuesTable_Procedure so sánh tên cột trong mảng `columns` với danh sách 11 issue fields hợp lệ, THE WriteRepoIssuesTable_Procedure SHALL thực hiện so sánh case-insensitive (ví dụ `"number"`, `"NUMBER"` và `"Number"` đều khớp issue field `Number`).
6. WHEN WriteRepoIssuesTable_Procedure so sánh tên cột trong mảng `columns` với project fields dạng `"[ProjectTitle] FieldName"`, THE WriteRepoIssuesTable_Procedure SHALL thực hiện so sánh case-insensitive trên phần `FieldName` và giữ nguyên `ProjectTitle` (ví dụ `"[My Board] status"` khớp với project field có ProjectTitle=`"My Board"` và FieldName=`"Status"`).
7. IF mảng `columns` chứa tên cột không hợp lệ (không phải một trong 11 issue fields hợp lệ VÀ không khớp pattern `"[...]..."` cho project field), THEN THE WriteRepoIssuesTable_Procedure SHALL bỏ qua tên cột đó — không raise error, không tạo cột cho tên không hợp lệ đó.
8. WHEN `WriteRepoIssuesTable` được gọi với `targetRange` hợp lệ và `GetRepoIssues` trả về `Success = True`, THE WriteRepoIssuesTable_Procedure SHALL ghi hàng tiêu đề cột tại hàng đầu tiên của vùng đích theo đúng danh sách cột đã xác định (từ tham số `columns` hoặc mặc định).
9. WHEN `WriteRepoIssuesTable` ghi dữ liệu cho một ô tại giao giữa hàng issue và cột project field, THE WriteRepoIssuesTable_Procedure SHALL parse chuỗi `RepoIssue.ProjectFields` (JSON array) để tìm giá trị tương ứng với `ProjectTitle` và `FieldName` của cột đó, ghi giá trị `value` nếu tìm thấy và ghi chuỗi rỗng nếu không tìm thấy.
10. WHEN `WriteRepoIssuesTable` ghi dữ liệu, THE WriteRepoIssuesTable_Procedure SHALL ghi mỗi đối tượng `RepoIssue` trên một hàng riêng biệt theo đúng thứ tự cột tiêu đề và đúng thứ tự xuất hiện trong mảng trả về.
11. WHEN `WriteRepoIssuesTable` ghi giá trị `CreatedAt`, `UpdatedAt` và `ClosedAt`, THE WriteRepoIssuesTable_Procedure SHALL ghi giá trị kiểu `Date` của VBA và áp dụng định dạng số `"yyyy-mm-dd hh:mm:ss"` cho các cột này, ghi chuỗi rỗng nếu `ClosedAt` là `Empty`.
12. WHEN `WriteRepoIssuesTable` hoàn tất việc ghi dữ liệu, THE WriteRepoIssuesTable_Procedure SHALL chuyển vùng đã ghi (gồm hàng tiêu đề và các hàng dữ liệu) thành một `ListObject` (bảng Excel) với tên `"tblRepoIssues_" & githubId & "_" & repoName` (sau khi thay thế các ký tự không hợp lệ trong tên `ListObject` bằng `_`) nếu vùng đích chưa thuộc một `ListObject` nào.
13. IF vùng đích đã thuộc một `ListObject` đang tồn tại, THEN THE WriteRepoIssuesTable_Procedure SHALL ghi dữ liệu vào trực tiếp `ListObject` đó (xóa các hàng dữ liệu cũ trong `ListObject` trước khi thêm hàng dữ liệu mới) và SHALL không tạo `ListObject` mới.
14. WHEN số lượng issue trả về bằng `0` và `Success = True`, THE WriteRepoIssuesTable_Procedure SHALL ghi hàng tiêu đề cột (theo danh sách cột đã xác định, nếu dùng mặc định thì chỉ ghi 11 issue fields vì không có dữ liệu để thu thập project fields) và SHALL không thêm hàng dữ liệu nào.
15. IF `GetRepoIssues` trả về `Success = False`, THEN THE WriteRepoIssuesTable_Procedure SHALL không ghi dữ liệu vào worksheet và SHALL phát sinh giá trị Err với `Err.Number = vbObjectError + 1001`, `Err.Source = "WriteRepoIssuesTable"` và `Err.Description` chứa giá trị `ErrorCode` và `ErrorMessage` từ kết quả của GetRepoIssues_Function.
16. IF `targetRange` là `Nothing`, THEN THE WriteRepoIssuesTable_Procedure SHALL phát sinh giá trị Err với `Err.Number = vbObjectError + 1002`, `Err.Source = "WriteRepoIssuesTable"` và `Err.Description` mô tả tham số `targetRange` không hợp lệ.
17. WHEN `WriteRepoIssuesTable` ghi dữ liệu, THE WriteRepoIssuesTable_Procedure SHALL áp dụng `AutoFit` chiều rộng cột cho toàn bộ các cột của bảng đã ghi.

### Requirement 20: Xóa code cũ liên quan đến GetRepoProjects

**User Story:** Là một người phát triển, tôi muốn xóa hoàn toàn code cũ liên quan đến GetRepoProjects để giữ codebase gọn gàng và tránh nhầm lẫn.

#### Acceptance Criteria

1. THE GitHub_API_Client SHALL không chứa hàm `GetRepoProjects`, `GetRepoProjectField`, thủ tục `WriteRepoProjectsTable`, hoặc thủ tục `SetMaxProjects`.
2. THE GitHub_API_Client SHALL không chứa kiểu dữ liệu `RepoProject` hoặc `RepoProjectsResult`.
3. THE GitHub_API_Client SHALL không chứa biến module-level `m_MaxProjects` hoặc hằng số liên quan đến `MaxProjects` (ví dụ `DEFAULT_MAX_PROJECTS`, `MIN_MAX_PROJECTS`, `MAX_MAX_PROJECTS`).
4. THE README.md SHALL không tham chiếu đến `GetRepoProjects`, `GetRepoProjectField`, `WriteRepoProjectsTable`, `SetMaxProjects`, `RepoProject`, hoặc `RepoProjectsResult`.
