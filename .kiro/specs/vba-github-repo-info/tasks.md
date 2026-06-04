# Implementation Plan: repo-issues-with-project-fields

## Overview

Replace the existing `GetRepoProjects`-related code in `GitHub_API_Client.bas` with new `GetRepoIssues`-related code, then update `README.md`. The existing infrastructure (validation, HTTP transport, response dispatch, JSON parsing utilities, `GetRepoInfo`) remains unchanged. All new code is VBA in a single `.bas` module.

## Tasks

- [x] 1. Remove old GetRepoProjects code and add new UDTs/constants
  - [x] 1.1 Remove GetRepoProjects-related code from GitHub_API_Client.bas
    - Delete `RepoProject` UDT, `RepoProjectsResult` UDT
    - Delete `GetRepoProjects` function, `GetRepoProjectField` function
    - Delete `WriteRepoProjectsTable` sub, `SetMaxProjects` sub
    - Delete `BuildRepoProjectsQuery` function, `ParseProjectsNodes` function
    - Delete `m_MaxProjects` module-level variable
    - Delete related constants (`DEFAULT_MAX_PROJECTS`, `MIN_MAX_PROJECTS`, `MAX_MAX_PROJECTS`, `ERR_BAD_MAX_PROJECTS`)
    - _Requirements: 20.1, 20.2, 20.3_

  - [x] 1.2 Add RepoIssue UDT, RepoIssuesResult UDT, and new constants/state
    - Add `RepoIssue` UDT with fields: Number, Title, Url, State, Author, CreatedAt, UpdatedAt, ClosedAt, Labels, Assignees, Milestone, ProjectFields
    - Add `RepoIssuesResult` UDT with fields: Items() As RepoIssue, TotalCount, Success, ErrorCode, ErrorMessage
    - Add `m_MaxIssues As Long` module-level variable
    - Add constants: `DEFAULT_MAX_ISSUES = 100`, `MIN_MAX_ISSUES = 1`, `MAX_MAX_ISSUES = 100`, `ERR_BAD_MAX_ISSUES = 1003`
    - _Requirements: 14.1, 14.2, 14.5, 14.6, 14.7_

- [x] 2. Implement SetMaxIssues and update EnsureDefaults
  - [x] 2.1 Implement SetMaxIssues sub and update EnsureDefaults
    - Add `Public Sub SetMaxIssues(ByVal maxIssues As Long)` that validates range [1, 100] and raises error `vbObjectError + ERR_BAD_MAX_ISSUES` if out of range
    - Update `EnsureDefaults` to set `m_MaxIssues = DEFAULT_MAX_ISSUES` if it's 0
    - _Requirements: 14.5, 14.6, 14.7_

- [x] 3. Implement BuildRepoIssuesQuery and GetRepoIssues
  - [x] 3.1 Implement BuildRepoIssuesQuery function
    - Create `Private Function BuildRepoIssuesQuery(ByVal states As String) As String`
    - Build GraphQL query with `repository.issues(first: $first, states: [...], orderBy: {field: CREATED_AT, direction: DESC})`
    - Include fields: number, title, url, state, author{login}, createdAt, updatedAt, closedAt, labels(first:10){nodes{name}}, assignees(first:5){nodes{login}}, milestone{title}
    - Include `projectItems(first: 10)` with `fieldValues(first: 20)` containing inline fragments for SingleSelectValue, TextValue, NumberValue, DateValue, IterationValue
    - Map states: "OPEN" → `[OPEN]`, "CLOSED" → `[CLOSED]`, "ALL" → `[OPEN, CLOSED]`
    - _Requirements: 15.1, 15.2, 15.3, 15.4, 15.5, 15.6_

  - [x] 3.2 Implement GetRepoIssues function
    - Create `Public Function GetRepoIssues(ByVal githubId As String, ByVal repoName As String, Optional ByVal states As String = "ALL") As Variant`
    - Flow: EnsureDefaults → check token → ValidateGithubId → ValidateRepoName → normalize/validate states → BuildRepoIssuesQuery → BuildRequestBody (with m_MaxIssues as first) → SendGraphQLRequest → ClassifyHttpResponse → parse JSON → ClassifyGraphQLErrors → ParseIssuesNodes → return RepoIssuesResult
    - Return failure RepoIssuesResult (empty Items, Success=False) for each error case with appropriate ErrorCode
    - States validation: UCase$(Trim$(states)) must be "OPEN", "CLOSED", or "ALL"; otherwise return INVALID_STATES
    - _Requirements: 14.1, 14.2, 14.3, 14.4, 14.8, 14.9, 14.10, 14.11, 15.7, 15.8, 17.1, 17.2, 17.3, 17.4, 17.5, 17.6, 17.7, 17.8, 17.9, 17.10, 17.11, 17.12, 17.13, 17.14, 17.15, 17.16, 17.17, 17.18, 17.19, 17.20, 17.21_

- [x] 4. Implement ParseIssuesNodes
  - [x] 4.1 Implement ParseIssuesNodes function
    - Create `Private Function ParseIssuesNodes(ByVal parsed As Object, ByRef issues() As RepoIssue, ByRef issuesCount As Long) As String`
    - Navigate to `data.repository.issues.nodes` in parsed JSON Dictionary
    - For each node: extract Number, Title, Url, State, Author (login or empty), CreatedAt/UpdatedAt/ClosedAt (via ParseIso8601, Empty for null closedAt)
    - Join labels from `labels.nodes[].name` with ", "
    - Join assignees from `assignees.nodes[].login` with ", "
    - Extract milestone title or empty string
    - Extract projectItems and serialize ProjectFields as JSON array `[{"project":"...","field":"...","value":"..."}, ...]`; use "[]" when no project items
    - Handle field value types: SingleSelectValue→name, TextValue→text, NumberValue→number, DateValue→date, IterationValue→title
    - Return "" on success, "PARSE_ERROR" on failure
    - _Requirements: 16.1, 16.2, 16.3, 16.4, 16.5, 16.6, 16.7, 16.8, 16.9, 16.10, 16.11, 16.12, 16.13, 16.14, 16.15, 16.16, 16.17, 16.18, 16.19_

- [x] 5. Checkpoint
  - Ensure all code compiles without errors in VBA editor, ask the user if questions arise.

- [x] 6. Implement GetRepoIssueField and WriteRepoIssuesTable
  - [x] 6.1 Implement GetRepoIssueField function
    - Create `Public Function GetRepoIssueField(ByVal githubId As String, ByVal repoName As String, ByVal issueIndex As Long, ByVal fieldName As String) As Variant`
    - Call GetRepoIssues internally, cache or direct call
    - Validate: if Success=False → return CVErr(xlErrNA); if issueIndex < 1 or > TotalCount → return CVErr(xlErrNA)
    - Match fieldName (case-insensitive) to RepoIssue fields: Number, Title, Url, State, Author, CreatedAt, UpdatedAt, ClosedAt, Labels, Assignees, Milestone, ProjectFields
    - If fieldName invalid → return CVErr(xlErrValue)
    - _Requirements: 18.1, 18.2, 18.3, 18.4, 18.5, 18.6_

  - [x] 6.2 Implement WriteRepoIssuesTable sub
    - Create `Public Sub WriteRepoIssuesTable(ByVal githubId As String, ByVal repoName As String, ByVal targetRange As Range, Optional ByVal columns As Variant)`
    - Check targetRange Is Nothing → Err.Raise vbObjectError + 1002
    - Call GetRepoIssues → if Success=False → Err.Raise vbObjectError + 1001 with ErrorCode + ErrorMessage
    - Resolve columns: if IsMissing/Empty → 11 default issue fields + discovered project fields (alpha sorted); else use provided array
    - Column matching: case-insensitive for issue fields; project fields format "[ProjectTitle] FieldName" with case-insensitive FieldName; skip invalid names
    - Write header row at targetRange, write data rows (issue fields direct from UDT, project fields by parsing ProjectFields JSON)
    - Date columns (CreatedAt, UpdatedAt, ClosedAt): write VBA Date, format "yyyy-mm-dd hh:mm:ss", empty string for Empty ClosedAt
    - Create/update ListObject: if targetRange already in ListObject → clear old data + write new; else create new named "tblRepoIssues_{githubId}_{repoName}" (sanitized via SafeListObjectName)
    - AutoFit all columns
    - If 0 issues + Success=True → write header only (default: 11 issue fields since no data to discover project fields)
    - _Requirements: 19.1, 19.2, 19.3, 19.4, 19.5, 19.6, 19.7, 19.8, 19.9, 19.10, 19.11, 19.12, 19.13, 19.14, 19.15, 19.16, 19.17_

- [x] 7. Checkpoint
  - Ensure all code compiles without errors in VBA editor, ask the user if questions arise.

- [x] 8. Update README.md
  - [x] 8.1 Update README.md documentation
    - Remove all references to GetRepoProjects, GetRepoProjectField, WriteRepoProjectsTable, SetMaxProjects, RepoProject, RepoProjectsResult
    - Add documentation for GetRepoIssues function (signature, parameters, return type, states filter)
    - Add documentation for GetRepoIssueField function (Excel UDF usage)
    - Add documentation for WriteRepoIssuesTable procedure (columns parameter explanation)
    - Add documentation for SetMaxIssues (range [1, 100], default 100)
    - Add example calling GetRepoIssues from VBA Sub
    - Add example calling GetRepoIssueField from Excel cell
    - Document ProjectFields JSON structure: `[{"project":"...","field":"...","value":"..."}, ...]` and how to parse in VBA
    - Note that token needs `read:project` scope for project fields and ProjectsCount
    - Include all ErrorCode values and their meanings
    - _Requirements: 13.4, 13.5, 13.6, 13.7, 20.4_

  - [ ]* 8.2 Write unit tests for GetRepoIssues with mock HTTP responses
    - Test successful response parsing
    - Test states parameter validation (OPEN, CLOSED, ALL, invalid)
    - Test error cases (MISSING_TOKEN, INVALID_GITHUB_ID, NOT_FOUND, etc.)
    - Test ParseIssuesNodes with various JSON structures
    - _Requirements: 14.8, 14.9, 14.10, 14.11, 16.1-16.19, 17.1-17.21_

  - [ ]* 8.3 Write unit tests for GetRepoIssueField and WriteRepoIssuesTable
    - Test GetRepoIssueField with valid/invalid fieldName and issueIndex
    - Test WriteRepoIssuesTable column resolution (default, custom, mixed issue+project fields)
    - Test error raising for Nothing targetRange and failed GetRepoIssues
    - _Requirements: 18.1-18.6, 19.1-19.17_

- [x] 9. Final checkpoint
  - Ensure all code compiles without errors, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- Each task references specific requirements for traceability
- Existing infrastructure (ValidateGithubId, ValidateRepoName, BuildRepoInfoQuery, BuildRequestBody, JsonEscape, SendGraphQLRequest, CreateHttpClient, ClassifyHttpResponse, ClassifyGraphQLErrors, ParseRepository, ParseIso8601, FormatUnixDate, SafeListObjectName, NullToEmpty, SafeHeader, GetRepoInfo, GetRepoField, SetGitHubToken, EnsureDefaults, MakeFailureRepoInfo, SetHttpFactoryForTest) is kept unchanged
- Implementation language: VBA (`.bas` module)
- No property-based tests included per user preference

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["1.1"] },
    { "id": 1, "tasks": ["1.2"] },
    { "id": 2, "tasks": ["2.1"] },
    { "id": 3, "tasks": ["3.1"] },
    { "id": 4, "tasks": ["3.2", "4.1"] },
    { "id": 5, "tasks": ["6.1", "6.2"] },
    { "id": 6, "tasks": ["8.1"] },
    { "id": 7, "tasks": ["8.2", "8.3"] }
  ]
}
```
