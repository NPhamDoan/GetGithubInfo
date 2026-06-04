Attribute VB_Name = "GitHub_API_Client"
Option Explicit

' ---------------------------------------------------------------------------
' User-Defined Types (UDTs)
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

Public Type RepoIssue
    Number          As Long
    Title           As String
    Url             As String
    State           As String
    Author          As String
    CreatedAt       As Date
    UpdatedAt       As Date
    ClosedAt        As Date
    Labels          As String
    Assignees       As String
    Milestone       As String
    ProjectFields   As String
End Type

Public Type RepoIssuesResult
    TotalCount      As Long
    Success         As Boolean
    ErrorCode       As String
    ErrorMessage    As String
End Type

' ---------------------------------------------------------------------------
' Module State
' ---------------------------------------------------------------------------
Private m_Token As String
Private m_MaxIssues As Long
Private m_HttpFactory As Object         ' Test seam for HTTP client injection
Private m_LastIssuesResult As RepoIssuesResult  ' Cache for last GetRepoIssues call
Private m_Issues() As RepoIssue         ' Array of issues from last GetRepoIssues call
Private m_IssuesCount As Long           ' Count of issues in m_Issues

' ---------------------------------------------------------------------------
' Constants
' ---------------------------------------------------------------------------
Private Const HTTP_TIMEOUT_MS As Long = 30000
Private Const GRAPHQL_URL As String = "https://api.github.com/graphql"
Private Const DEFAULT_USER_AGENT As String = "VBA-GitHub-GraphQL-Client"
Private Const DEFAULT_MAX_ISSUES As Long = 100
Private Const MIN_MAX_ISSUES As Long = 1
Private Const MAX_MAX_ISSUES As Long = 100

' Error numbers (to be used with vbObjectError + ERR_xxx)
Private Const ERR_WRITE_FAIL As Long = 1001
Private Const ERR_NULL_RANGE As Long = 1002
Private Const ERR_BAD_MAX_ISSUES As Long = 1003

' ---------------------------------------------------------------------------
' EnsureDefaults - Initialize module-level defaults on first use
' ---------------------------------------------------------------------------
Public Sub EnsureDefaults()
    If m_MaxIssues = 0 Then m_MaxIssues = DEFAULT_MAX_ISSUES
End Sub

' ---------------------------------------------------------------------------
' SetMaxIssues - Configure the maximum number of issues to retrieve
' Validates maxIssues >= 1; raises vbObjectError + ERR_BAD_MAX_ISSUES if
' out of range. Values > 100 trigger automatic pagination in GetRepoIssues.
' Req 14.5, 14.6, 14.7
' ---------------------------------------------------------------------------
Public Sub SetMaxIssues(ByVal maxIssues As Long)
    If maxIssues < MIN_MAX_ISSUES Then
        Err.Raise vbObjectError + ERR_BAD_MAX_ISSUES, "SetMaxIssues", _
                  "maxIssues must be >= 1"
        Exit Sub
    End If
    m_MaxIssues = maxIssues
End Sub

' ---------------------------------------------------------------------------
' Validation (pure)
' ---------------------------------------------------------------------------

' ValidateGithubId
' Returns True if s is a valid GitHub username/org:
'   - Not empty/whitespace after Trim$
'   - Length <= 39
'   - Only characters in [A-Za-z0-9-]
' Req 3.1, 3.3, 15.2, 15.4
Private Function ValidateGithubId(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then Exit Function
    If Len(s) > 39 Then Exit Function
    Dim i As Long, ch As Long
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or _
                (ch >= 65 And ch <= 90) Or _
                (ch >= 97 And ch <= 122) Or _
                ch = 45) Then Exit Function
    Next i
    ValidateGithubId = True
End Function

' ValidateRepoName
' Returns True if s is a valid GitHub repository name:
'   - Not empty/whitespace after Trim$
'   - Length <= 100
'   - Only characters in [A-Za-z0-9._-]
' Req 3.2, 3.4, 15.3, 15.5
Private Function ValidateRepoName(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then Exit Function
    If Len(s) > 100 Then Exit Function
    Dim i As Long, ch As Long
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or _
                (ch >= 65 And ch <= 90) Or _
                (ch >= 97 And ch <= 122) Or _
                ch = 45 Or _
                ch = 46 Or _
                ch = 95) Then Exit Function
    Next i
    ValidateRepoName = True
End Function

' ---------------------------------------------------------------------------
' JsonEscape - Escape a string for safe use as a JSON string value
' Handles: " \ / (optional) backspace formfeed newline return tab
'          and all control characters U+0000..U+001F as \u00XX
' Req 4.4, 16.4
' ---------------------------------------------------------------------------
Private Function JsonEscape(ByVal s As String) As String
    Dim i As Long
    Dim ch As Long
    Dim result As String
    
    result = ""
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        Select Case ch
            Case 34 ' "
                result = result & "\"""
            Case 92 ' \
                result = result & "\\"
            Case 8  ' backspace
                result = result & "\b"
            Case 12 ' formfeed
                result = result & "\f"
            Case 10 ' newline
                result = result & "\n"
            Case 13 ' carriage return
                result = result & "\r"
            Case 9  ' tab
                result = result & "\t"
            Case 0 To 31 ' other control characters
                result = result & "\u" & Right$("0000" & Hex$(ch), 4)
            Case Else
                result = result & ChrW$(ch)
        End Select
    Next i
    
    JsonEscape = result
End Function

' ---------------------------------------------------------------------------
' BuildRepoInfoQuery - Returns the GraphQL query string for GetRepoInfo
' Uses parameterized variables $owner and $name (no string interpolation)
' Req 4.1, 4.2
' ---------------------------------------------------------------------------
Private Function BuildRepoInfoQuery() As String
    BuildRepoInfoQuery = "query($owner: String!, $name: String!) { repository(owner: $owner, name: $name) { nameWithOwner description url stargazerCount forkCount isPrivate isFork isArchived diskUsage createdAt updatedAt defaultBranchRef { name } primaryLanguage { name } issues(states: OPEN) { totalCount } projectsV2 { totalCount } } }"
End Function

' ---------------------------------------------------------------------------
' BuildRepoIssuesQuery - Returns the GraphQL query string for GetRepoIssues
' Accepts a states parameter ("OPEN", "CLOSED", or "ALL") and embeds the
' corresponding enum filter in the query. Uses parameterized variables
' $owner, $name, $first for user-supplied values.
' Req 15.1, 15.2, 15.3, 15.4, 15.5, 15.6
' ---------------------------------------------------------------------------
Private Function BuildRepoIssuesQuery(ByVal states As String) As String
    Dim statesFilter As String
    Select Case states
        Case "OPEN"
            statesFilter = "[OPEN]"
        Case "CLOSED"
            statesFilter = "[CLOSED]"
        Case Else ' "ALL"
            statesFilter = "[OPEN, CLOSED]"
    End Select
    
    BuildRepoIssuesQuery = "query($owner: String!, $name: String!, $first: Int!, $after: String) { " & _
        "repository(owner: $owner, name: $name) { " & _
        "issues(first: $first, after: $after, states: " & statesFilter & ", orderBy: {field: CREATED_AT, direction: DESC}) { " & _
        "pageInfo { hasNextPage endCursor } " & _
        "totalCount nodes { " & _
        "number title url state author { login } createdAt updatedAt closedAt " & _
        "labels(first: 10) { nodes { name } } " & _
        "assignees(first: 5) { nodes { login } } " & _
        "milestone { title } " & _
        "projectItems(first: 10) { nodes { project { title } " & _
        "fieldValues(first: 20) { nodes { " & _
        "... on ProjectV2ItemFieldSingleSelectValue { field { ... on ProjectV2SingleSelectField { name } } name } " & _
        "... on ProjectV2ItemFieldTextValue { field { ... on ProjectV2Field { name } } text } " & _
        "... on ProjectV2ItemFieldNumberValue { field { ... on ProjectV2Field { name } } number } " & _
        "... on ProjectV2ItemFieldDateValue { field { ... on ProjectV2Field { name } } date } " & _
        "... on ProjectV2ItemFieldIterationValue { field { ... on ProjectV2IterationField { name } } title } " & _
        "} } } } } } } }"
End Function

' ---------------------------------------------------------------------------
' BuildRequestBody - Builds the JSON request body for a GraphQL POST request
' Combines the query string and variables (owner, name, optional first) into
' a single-line JSON object. Uses JsonEscape for safe string embedding.
' The "first" key is only included when first >= 0.
' Req 4.3, 4.4, 16.3, 16.4
' ---------------------------------------------------------------------------
Private Function BuildRequestBody(ByVal query As String, ByVal owner As String, ByVal name As String, Optional ByVal first As Long = -1, Optional ByVal after As String = "") As String
    Dim body As String
    body = "{""query"":""" & JsonEscape(query) & """,""variables"":{""owner"":""" & JsonEscape(owner) & """,""name"":""" & JsonEscape(name) & """"
    
    If first >= 0 Then
        body = body & ",""first"":" & CStr(first)
    End If
    
    If after <> "" Then
        body = body & ",""after"":""" & JsonEscape(after) & """"
    End If
    
    body = body & "}}"
    BuildRequestBody = body
End Function

' ---------------------------------------------------------------------------
' Utilities (pure)
' ---------------------------------------------------------------------------

' ParseIso8601 - Parse ISO 8601 date string to VBA Date (UTC)
' Accepts: YYYY-MM-DDTHH:MM:SSZ or YYYY-MM-DDTHH:MM:SS±HH:MM
' Returns 0 on any parse error
' Req 5.10, 5.11
Private Function ParseIso8601(ByVal s As String) As Date
    On Error Resume Next
    Dim yr As Long, mo As Long, dy As Long
    Dim hr As Long, mn As Long, sc As Long
    Dim offsetSign As Long, offsetHr As Long, offsetMn As Long
    Dim tPos As Long
    
    ' Find T separator
    tPos = InStr(1, s, "T", vbTextCompare)
    If tPos = 0 Then
        ParseIso8601 = 0
        Exit Function
    End If
    
    ' Parse date part: YYYY-MM-DD
    yr = CLng(Mid$(s, 1, 4))
    mo = CLng(Mid$(s, 6, 2))
    dy = CLng(Mid$(s, 9, 2))
    
    ' Parse time part: HH:MM:SS
    hr = CLng(Mid$(s, tPos + 1, 2))
    mn = CLng(Mid$(s, tPos + 4, 2))
    sc = CLng(Mid$(s, tPos + 7, 2))
    
    ' Parse timezone offset
    Dim tzPart As String
    tzPart = Mid$(s, tPos + 9)
    offsetSign = 0
    offsetHr = 0
    offsetMn = 0
    
    If Len(tzPart) > 0 Then
        If Left$(tzPart, 1) = "Z" Then
            ' UTC, no offset
        ElseIf Left$(tzPart, 1) = "+" Then
            offsetSign = -1  ' subtract offset to get UTC
            offsetHr = CLng(Mid$(tzPart, 2, 2))
            offsetMn = CLng(Mid$(tzPart, 5, 2))
        ElseIf Left$(tzPart, 1) = "-" Then
            offsetSign = 1   ' add offset to get UTC
            offsetHr = CLng(Mid$(tzPart, 2, 2))
            offsetMn = CLng(Mid$(tzPart, 5, 2))
        End If
    End If
    
    ' Build date
    Dim result As Date
    result = DateSerial(yr, mo, dy) + TimeSerial(hr, mn, sc)
    
    ' Apply timezone offset to convert to UTC
    If offsetSign <> 0 Then
        result = DateAdd("h", offsetSign * offsetHr, result)
        result = DateAdd("n", offsetSign * offsetMn, result)
    End If
    
    If Err.Number <> 0 Then
        ParseIso8601 = 0
    Else
        ParseIso8601 = result
    End If
    On Error GoTo 0
End Function

' FormatUnixDate - Convert Unix timestamp to local date/time string
' Req 7.2, 19.7
Private Function FormatUnixDate(ByVal unixSeconds As Double) As String
    Dim d As Date
    d = #1/1/1970# + (unixSeconds / 86400#)
    FormatUnixDate = Format$(d, "yyyy-mm-dd hh:mm:ss")
End Function

' SafeListObjectName - Sanitize name for use as Excel ListObject name
' Replace any char not in [A-Za-z0-9_] with _, max 255, no leading digit
' Req 21.6
Private Function SafeListObjectName(ByVal raw As String) As String
    Dim i As Long, ch As Long
    Dim result As String
    
    result = ""
    For i = 1 To Len(raw)
        ch = AscW(Mid$(raw, i, 1))
        If (ch >= 48 And ch <= 57) Or _
           (ch >= 65 And ch <= 90) Or _
           (ch >= 97 And ch <= 122) Or _
           ch = 95 Then
            result = result & ChrW$(ch)
        Else
            result = result & "_"
        End If
    Next i
    
    ' Ensure first character is not a digit
    If Len(result) > 0 Then
        ch = AscW(Left$(result, 1))
        If ch >= 48 And ch <= 57 Then
            result = "_" & result
        End If
    End If
    
    ' Truncate to 255 characters
    If Len(result) > 255 Then
        result = Left$(result, 255)
    End If
    
    SafeListObjectName = result
End Function

' NullToEmpty - Convert Null/Nothing to empty string
' Req 17.8, 17.9
Private Function NullToEmpty(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Then
        NullToEmpty = ""
    ElseIf IsObject(v) Then
        If v Is Nothing Then
            NullToEmpty = ""
        Else
            NullToEmpty = CStr(v)
        End If
    Else
        NullToEmpty = CStr(v)
    End If
    On Error GoTo 0
End Function

' SafeHeader - Safely read HTTP response header
' Returns empty string if header not found or error occurs
Private Function SafeHeader(ByVal http As Object, ByVal headerName As String) As String
    On Error Resume Next
    SafeHeader = http.getResponseHeader(headerName)
    If Err.Number <> 0 Then SafeHeader = ""
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------
' ClassifyHttpResponse - Classify HTTP response into error category
' Returns one of: "OK", "UNAUTHORIZED", "FORBIDDEN", "RATE_LIMITED",
'                 "SERVER_ERROR", "HTTP_ERROR"
' Sets errCode and errMsg ByRef for non-OK results.
' Req 7.1, 7.2, 8.1, 8.2, 8.3, 8.4, 19.6, 19.7, 19.8, 19.9, 19.10
' ---------------------------------------------------------------------------
Private Function ClassifyHttpResponse(ByVal httpStatus As Long, _
                                      ByVal rateLimitRemaining As String, _
                                      ByVal rateLimitReset As String, _
                                      ByVal body As String, _
                                      ByRef errCode As String, _
                                      ByRef errMsg As String) As String
    Select Case httpStatus
        Case 200
            ClassifyHttpResponse = "OK"
            
        Case 401
            errCode = "UNAUTHORIZED"
            errMsg = "HTTP 401 Unauthorized"
            ClassifyHttpResponse = "UNAUTHORIZED"
            
        Case 403
            If rateLimitRemaining = "0" Then
                errCode = "RATE_LIMITED"
                errMsg = "Rate limit reset at " & FormatUnixDate(CDbl(rateLimitReset))
                ClassifyHttpResponse = "RATE_LIMITED"
            Else
                errCode = "FORBIDDEN"
                errMsg = "HTTP 403 Forbidden"
                ClassifyHttpResponse = "FORBIDDEN"
            End If
            
        Case 500 To 599
            errCode = "SERVER_ERROR"
            errMsg = "HTTP " & CStr(httpStatus) & " Server Error"
            ClassifyHttpResponse = "SERVER_ERROR"
            
        Case Else
            errCode = "HTTP_ERROR"
            errMsg = "HTTP " & CStr(httpStatus)
            ClassifyHttpResponse = "HTTP_ERROR"
    End Select
End Function

' ---------------------------------------------------------------------------
' ClassifyGraphQLErrors - Classify GraphQL-level errors from parsed response
' Returns True if a GraphQL error was found (errCode/errMsg set), False if OK.
' Decision table:
'   errors[0].type = "NOT_FOUND"   → errCode="NOT_FOUND", errMsg=githubId/repoName
'   errors[0].type = "FORBIDDEN"   → errCode="FORBIDDEN", errMsg=errors[0].message
'   errors[0].type = "RATE_LIMITED" → errCode="RATE_LIMITED", errMsg=errors[0].message
'   Other type                      → errCode="GRAPHQL_ERROR", errMsg=errors[0].message
'   errors empty/absent AND data.repository=Nothing → errCode="NOT_FOUND"
'   errors empty/absent AND data.repository exists  → return False (no error)
' Req 6.1, 6.2, 6.3, 6.4, 6.5, 19.1, 19.2, 19.3, 19.4, 19.5
' ---------------------------------------------------------------------------
Private Function ClassifyGraphQLErrors(ByVal parsed As Object, _
                                       ByVal githubId As String, _
                                       ByVal repoName As String, _
                                       ByRef errCode As String, _
                                       ByRef errMsg As String) As Boolean
    Dim hasErrors As Boolean
    Dim errorsCol As Object
    Dim errCount As Long
    Dim firstErr As Object
    Dim errType As String
    
    ' Check if "errors" key exists and has items
    hasErrors = False
    On Error Resume Next
    If parsed.Exists("errors") Then
        Set errorsCol = parsed("errors")
        errCount = errorsCol.Count
        If errCount > 0 Then
            hasErrors = True
        End If
    End If
    On Error GoTo 0
    
    If hasErrors Then
        ' Read first error's type and message
        On Error Resume Next
        Set firstErr = errorsCol(1)  ' Collection is 1-based in VBA (JsonConverter)
        errType = ""
        errType = firstErr("type")
        If Err.Number <> 0 Then errType = ""
        Err.Clear
        
        Dim errMessage As String
        errMessage = ""
        errMessage = firstErr("message")
        If Err.Number <> 0 Then errMessage = ""
        On Error GoTo 0
        
        ' Map error type to error code
        Select Case errType
            Case "NOT_FOUND"
                errCode = "NOT_FOUND"
                errMsg = githubId & "/" & repoName
            Case "FORBIDDEN"
                errCode = "FORBIDDEN"
                errMsg = errMessage
            Case "RATE_LIMITED"
                errCode = "RATE_LIMITED"
                errMsg = errMessage
            Case Else
                errCode = "GRAPHQL_ERROR"
                errMsg = errMessage
        End Select
        
        ClassifyGraphQLErrors = True
        Exit Function
    End If
    
    ' No errors array or empty — check if data.repository is Nothing
    Dim repo As Object
    On Error Resume Next
    Set repo = parsed("data")("repository")
    If Err.Number <> 0 Then
        ' Cannot access data.repository — treat as NOT_FOUND
        Err.Clear
        On Error GoTo 0
        errCode = "NOT_FOUND"
        errMsg = githubId & "/" & repoName
        ClassifyGraphQLErrors = True
        Exit Function
    End If
    On Error GoTo 0
    
    If repo Is Nothing Then
        errCode = "NOT_FOUND"
        errMsg = githubId & "/" & repoName
        ClassifyGraphQLErrors = True
        Exit Function
    End If
    
    ' No error found — repository data exists
    ClassifyGraphQLErrors = False
End Function

' ---------------------------------------------------------------------------
' ParseRepository - Parse repository data from GraphQL response into RepoInfo
' Returns "" on success, "PARSE_ERROR" on failure.
' Req 5.1-5.17, 10.2, 10.3
' ---------------------------------------------------------------------------
Private Function ParseRepository(ByVal parsed As Object, _
                                 ByVal githubId As String, _
                                 ByVal repoName As String, _
                                 ByRef result As RepoInfo) As String
    Dim repo As Object
    Dim tempObj As Object
    
    On Error Resume Next
    
    ' Access data.repository
    Set repo = parsed("data")("repository")
    If Err.Number <> 0 Then
        On Error GoTo 0
        ParseRepository = "PARSE_ERROR"
        Exit Function
    End If
    On Error GoTo 0
    
    If repo Is Nothing Then
        ParseRepository = "PARSE_ERROR"
        Exit Function
    End If
    
    On Error Resume Next
    
    ' --- Mandatory fields ---
    ' nameWithOwner (Req 5.2, 10.3)
    result.NameWithOwner = CStr(repo("nameWithOwner"))
    If Err.Number <> 0 Then
        On Error GoTo 0
        ParseRepository = "PARSE_ERROR"
        Exit Function
    End If
    
    ' createdAt (Req 5.10, 10.3)
    Dim createdAtStr As String
    createdAtStr = CStr(repo("createdAt"))
    If Err.Number <> 0 Then
        On Error GoTo 0
        ParseRepository = "PARSE_ERROR"
        Exit Function
    End If
    
    ' updatedAt (Req 5.11, 10.3)
    Dim updatedAtStr As String
    updatedAtStr = CStr(repo("updatedAt"))
    If Err.Number <> 0 Then
        On Error GoTo 0
        ParseRepository = "PARSE_ERROR"
        Exit Function
    End If
    
    On Error GoTo 0
    
    ' Parse dates
    result.CreatedAt = ParseIso8601(createdAtStr)
    result.UpdatedAt = ParseIso8601(updatedAtStr)
    
    On Error Resume Next
    
    ' --- Nullable string fields ---
    ' description (Req 5.3)
    result.Description = NullToEmpty(repo("description"))
    If Err.Number <> 0 Then
        Err.Clear
        result.Description = ""
    End If
    
    ' primaryLanguage (Req 5.4) — nullable object
    Set tempObj = Nothing
    Set tempObj = repo("primaryLanguage")
    If Err.Number <> 0 Then
        Err.Clear
        result.PrimaryLanguage = ""
    ElseIf tempObj Is Nothing Then
        result.PrimaryLanguage = ""
    Else
        result.PrimaryLanguage = CStr(tempObj("name"))
        If Err.Number <> 0 Then
            Err.Clear
            result.PrimaryLanguage = ""
        End If
    End If
    
    ' defaultBranchRef (Req 5.8) — nullable object
    Set tempObj = Nothing
    Set tempObj = repo("defaultBranchRef")
    If Err.Number <> 0 Then
        Err.Clear
        result.DefaultBranch = ""
    ElseIf tempObj Is Nothing Then
        result.DefaultBranch = ""
    Else
        result.DefaultBranch = CStr(tempObj("name"))
        If Err.Number <> 0 Then
            Err.Clear
            result.DefaultBranch = ""
        End If
    End If
    
    ' --- Numeric fields ---
    ' stargazerCount (Req 5.5)
    result.Stars = CLng(repo("stargazerCount"))
    If Err.Number <> 0 Then
        Err.Clear
        result.Stars = 0
    End If
    
    ' forkCount (Req 5.6)
    result.Forks = CLng(repo("forkCount"))
    If Err.Number <> 0 Then
        Err.Clear
        result.Forks = 0
    End If
    
    ' issues.totalCount → OpenIssues (Req 5.7)
    Set tempObj = Nothing
    Set tempObj = repo("issues")
    If Err.Number <> 0 Then
        Err.Clear
        result.OpenIssues = 0
    ElseIf tempObj Is Nothing Then
        result.OpenIssues = 0
    Else
        result.OpenIssues = CLng(tempObj("totalCount"))
        If Err.Number <> 0 Then
            Err.Clear
            result.OpenIssues = 0
        End If
    End If
    
    ' url (Req 5.9)
    result.Url = CStr(repo("url"))
    If Err.Number <> 0 Then
        Err.Clear
        result.Url = ""
    End If
    
    ' --- Boolean fields ---
    ' isPrivate (Req 5.12)
    result.IsPrivate = CBool(repo("isPrivate"))
    If Err.Number <> 0 Then
        Err.Clear
        result.IsPrivate = False
    End If
    
    ' isFork (Req 5.13)
    result.IsFork = CBool(repo("isFork"))
    If Err.Number <> 0 Then
        Err.Clear
        result.IsFork = False
    End If
    
    ' isArchived (Req 5.14)
    result.IsArchived = CBool(repo("isArchived"))
    If Err.Number <> 0 Then
        Err.Clear
        result.IsArchived = False
    End If
    
    ' --- Other numeric ---
    ' diskUsage (Req 5.15)
    result.DiskUsageKB = CLng(repo("diskUsage"))
    If Err.Number <> 0 Then
        Err.Clear
        result.DiskUsageKB = 0
    End If
    
    ' projectsV2.totalCount → ProjectsCount (Req 5.16) — nullable object
    Set tempObj = Nothing
    Set tempObj = repo("projectsV2")
    If Err.Number <> 0 Then
        Err.Clear
        result.ProjectsCount = 0
    ElseIf tempObj Is Nothing Then
        result.ProjectsCount = 0
    Else
        result.ProjectsCount = CLng(tempObj("totalCount"))
        If Err.Number <> 0 Then
            Err.Clear
            result.ProjectsCount = 0
        End If
    End If
    
    On Error GoTo 0
    
    ' --- Success (Req 5.17) ---
    result.Success = True
    result.ErrorCode = ""
    result.ErrorMessage = ""
    
    ParseRepository = ""
End Function

' ---------------------------------------------------------------------------
' ParseIssuesNodes - Parse issues nodes from GraphQL response into array
' Navigates to data.repository.issues.nodes and extracts each issue's fields
' including labels, assignees, milestone, and projectItems with field values.
' Returns "" on success, "PARSE_ERROR" on failure.
' Req 16.1-16.19
' ---------------------------------------------------------------------------
Private Function ParseIssuesNodes(ByVal parsed As Object, _
                                  ByRef issues() As RepoIssue, _
                                  ByRef issuesCount As Long) As String
    Dim nodes As Object
    Dim node As Object
    Dim i As Long
    Dim nodeCount As Long
    
    On Error GoTo ParseFail
    
    ' Navigate to data.repository.issues.nodes
    Set nodes = parsed("data")("repository")("issues")("nodes")
    nodeCount = nodes.Count
    
    ' Handle empty nodes (Req 16.16, 16.17)
    If nodeCount = 0 Then
        issuesCount = 0
        ReDim issues(0 To 0)
        ParseIssuesNodes = ""
        Exit Function
    End If
    
    ' Allocate array
    ReDim issues(1 To nodeCount)
    issuesCount = nodeCount
    
    ' Process each issue node
    Dim j As Long
    For i = 1 To nodeCount
        Set node = nodes(i)
        
        ' Basic fields (Req 16.2, 16.3, 16.4, 16.5)
        issues(i).Number = CLng(node("number"))
        issues(i).Title = CStr(node("title"))
        issues(i).Url = CStr(node("url"))
        issues(i).State = CStr(node("state"))
        
        ' Author (Req 16.6) - may be null
        Dim authorObj As Object
        Set authorObj = Nothing
        On Error Resume Next
        Set authorObj = node("author")
        Err.Clear
        On Error GoTo ParseFail
        If authorObj Is Nothing Then
            issues(i).Author = ""
        Else
            On Error Resume Next
            issues(i).Author = CStr(authorObj("login"))
            If Err.Number <> 0 Then
                Err.Clear
                issues(i).Author = ""
            End If
            On Error GoTo ParseFail
        End If
        
        ' CreatedAt (Req 16.7)
        issues(i).CreatedAt = ParseIso8601(CStr(node("createdAt")))
        
        ' UpdatedAt (Req 16.8)
        issues(i).UpdatedAt = ParseIso8601(CStr(node("updatedAt")))
        
        ' ClosedAt (Req 16.9, 16.10) - null means Empty (default Date)
        Dim closedAtVal As Variant
        On Error Resume Next
        closedAtVal = node("closedAt")
        On Error GoTo ParseFail
        If IsNull(closedAtVal) Then
            ' Leave ClosedAt as Empty (default Date value)
        Else
            issues(i).ClosedAt = ParseIso8601(CStr(closedAtVal))
        End If
        
        ' Labels (Req 16.11)
        Dim labelsNodes As Object
        Dim labelStr As String
        labelStr = ""
        On Error Resume Next
        Set labelsNodes = node("labels")("nodes")
        On Error GoTo ParseFail
        If Not labelsNodes Is Nothing Then
            If labelsNodes.Count > 0 Then
                For j = 1 To labelsNodes.Count
                    If j > 1 Then labelStr = labelStr & ", "
                    labelStr = labelStr & CStr(labelsNodes(j)("name"))
                Next j
            End If
        End If
        issues(i).Labels = labelStr
        
        ' Assignees (Req 16.12)
        Dim assigneesNodes As Object
        Dim assigneeStr As String
        assigneeStr = ""
        On Error Resume Next
        Set assigneesNodes = node("assignees")("nodes")
        On Error GoTo ParseFail
        If Not assigneesNodes Is Nothing Then
            If assigneesNodes.Count > 0 Then
                For j = 1 To assigneesNodes.Count
                    If j > 1 Then assigneeStr = assigneeStr & ", "
                    assigneeStr = assigneeStr & CStr(assigneesNodes(j)("login"))
                Next j
            End If
        End If
        issues(i).Assignees = assigneeStr
        
        ' Milestone (Req 16.13)
        Dim milestoneObj As Object
        Set milestoneObj = Nothing
        On Error Resume Next
        Set milestoneObj = node("milestone")
        Err.Clear
        On Error GoTo ParseFail
        If milestoneObj Is Nothing Then
            issues(i).Milestone = ""
        Else
            On Error Resume Next
            issues(i).Milestone = CStr(milestoneObj("title"))
            If Err.Number <> 0 Then
                Err.Clear
                issues(i).Milestone = ""
            End If
            On Error GoTo ParseFail
        End If
        
        ' ProjectFields (Req 16.14, 16.15)
        Dim projectItemsNodes As Object
        Set projectItemsNodes = Nothing
        On Error Resume Next
        Set projectItemsNodes = node("projectItems")("nodes")
        On Error GoTo ParseFail
        
        If projectItemsNodes Is Nothing Then
            issues(i).ProjectFields = "[]"
        ElseIf projectItemsNodes.Count = 0 Then
            issues(i).ProjectFields = "[]"
        Else
            ' Build JSON array of project field values
            Dim pfJson As String
            Dim pfFirst As Boolean
            pfJson = "["
            pfFirst = True
            
            Dim k As Long
            Dim piNode As Object
            Dim projTitle As String
            Dim fvNodes As Object
            Dim fvNode As Object
            Dim fieldName As String
            Dim fieldValue As String
            
            For k = 1 To projectItemsNodes.Count
                Set piNode = projectItemsNodes(k)
                
                ' Get project title
                On Error Resume Next
                projTitle = CStr(piNode("project")("title"))
                If Err.Number <> 0 Then
                    Err.Clear
                    projTitle = ""
                End If
                On Error GoTo ParseFail
                
                ' Get fieldValues.nodes
                Set fvNodes = Nothing
                On Error Resume Next
                Set fvNodes = piNode("fieldValues")("nodes")
                On Error GoTo ParseFail
                
                If Not fvNodes Is Nothing Then
                    Dim m As Long
                    For m = 1 To fvNodes.Count
                        Set fvNode = fvNodes(m)
                        fieldName = ""
                        fieldValue = ""
                        
                        ' Determine field value type by checking which keys exist
                        ' Detection strategy:
                        '   has "text" → TextValue
                        '   has "number" → NumberValue
                        '   has "date" → DateValue
                        '   has "title" → IterationValue
                        '   has "name" AND "field" → SingleSelectValue
                        '   else → skip
                        
                        Dim hasText As Boolean
                        Dim hasNumber As Boolean
                        Dim hasDate As Boolean
                        Dim hasTitle As Boolean
                        Dim hasName As Boolean
                        Dim hasField As Boolean
                        
                        hasText = False
                        hasNumber = False
                        hasDate = False
                        hasTitle = False
                        hasName = False
                        hasField = False
                        
                        On Error Resume Next
                        hasField = fvNode.Exists("field")
                        hasText = fvNode.Exists("text")
                        hasNumber = fvNode.Exists("number")
                        hasDate = fvNode.Exists("date")
                        hasTitle = fvNode.Exists("title")
                        hasName = fvNode.Exists("name")
                        On Error GoTo ParseFail
                        
                        If hasText Then
                            ' TextValue: field.name → fieldName, text → fieldValue
                            On Error Resume Next
                            fieldName = CStr(fvNode("field")("name"))
                            If Err.Number <> 0 Then Err.Clear: fieldName = ""
                            fieldValue = CStr(fvNode("text"))
                            If Err.Number <> 0 Then Err.Clear: fieldValue = ""
                            On Error GoTo ParseFail
                        ElseIf hasNumber Then
                            ' NumberValue: field.name → fieldName, number → fieldValue
                            On Error Resume Next
                            fieldName = CStr(fvNode("field")("name"))
                            If Err.Number <> 0 Then Err.Clear: fieldName = ""
                            fieldValue = CStr(fvNode("number"))
                            If Err.Number <> 0 Then Err.Clear: fieldValue = ""
                            On Error GoTo ParseFail
                        ElseIf hasDate Then
                            ' DateValue: field.name → fieldName, date → fieldValue
                            On Error Resume Next
                            fieldName = CStr(fvNode("field")("name"))
                            If Err.Number <> 0 Then Err.Clear: fieldName = ""
                            fieldValue = CStr(fvNode("date"))
                            If Err.Number <> 0 Then Err.Clear: fieldValue = ""
                            On Error GoTo ParseFail
                        ElseIf hasTitle Then
                            ' IterationValue: field.name → fieldName, title → fieldValue
                            On Error Resume Next
                            fieldName = CStr(fvNode("field")("name"))
                            If Err.Number <> 0 Then Err.Clear: fieldName = ""
                            fieldValue = CStr(fvNode("title"))
                            If Err.Number <> 0 Then Err.Clear: fieldValue = ""
                            On Error GoTo ParseFail
                        ElseIf hasName And hasField Then
                            ' SingleSelectValue: field.name → fieldName, name → fieldValue
                            On Error Resume Next
                            fieldName = CStr(fvNode("field")("name"))
                            If Err.Number <> 0 Then Err.Clear: fieldName = ""
                            fieldValue = CStr(fvNode("name"))
                            If Err.Number <> 0 Then Err.Clear: fieldValue = ""
                            On Error GoTo ParseFail
                        Else
                            ' Unknown type or empty fragment match - skip
                            GoTo NextFieldValue
                        End If
                        
                        ' Append to JSON array
                        If Not pfFirst Then pfJson = pfJson & ","
                        pfJson = pfJson & "{""project"":""" & JsonEscape(projTitle) & _
                                 """,""field"":""" & JsonEscape(fieldName) & _
                                 """,""value"":""" & JsonEscape(fieldValue) & """}"
                        pfFirst = False
                        
NextFieldValue:
                    Next m
                End If
            Next k
            
            pfJson = pfJson & "]"
            issues(i).ProjectFields = pfJson
        End If
    Next i
    
    ParseIssuesNodes = ""
    Exit Function

ParseFail:
    ParseIssuesNodes = "PARSE_ERROR"
End Function

' ---------------------------------------------------------------------------
' CreateHttpClient - Factory method for HTTP client (test seam)
' If m_HttpFactory is set (via SetHttpFactoryForTest), delegates to it;
' otherwise creates a real MSXML2.ServerXMLHTTP.6.0 instance.
' Req 11.3
' ---------------------------------------------------------------------------
Private Function CreateHttpClient() As Object
    If m_HttpFactory Is Nothing Then
        Set CreateHttpClient = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Else
        Set CreateHttpClient = m_HttpFactory.Create()
    End If
End Function

' ---------------------------------------------------------------------------
' SetHttpFactoryForTest - Test seam to inject a mock HTTP factory
' The factory object must expose a Create() method returning an object
' compatible with MSXML2.ServerXMLHTTP interface (Open, setRequestHeader,
' setTimeouts, send, Status, responseText, getResponseHeader).
' This sub is intentionally NOT documented in README.
' ---------------------------------------------------------------------------
Public Sub SetHttpFactoryForTest(ByVal factory As Object)
    Set m_HttpFactory = factory
End Sub

' ---------------------------------------------------------------------------
' SendGraphQLRequest - Send a GraphQL POST request to GitHub API
' Returns True on successful HTTP round-trip (caller inspects httpStatus).
' Returns False on network/timeout error (errCode and errMsg set).
' Uses CreateHttpClient for testability.
' Req 1.3, 1.4, 1.5, 1.6, 2.2, 9.1, 9.2, 9.3, 11.3, 14.3, 14.4,
'     19.11, 19.12, 19.13
' ---------------------------------------------------------------------------
Private Function SendGraphQLRequest(ByVal jsonBody As String, _
                                    ByRef httpStatus As Long, _
                                    ByRef rateLimitRemaining As String, _
                                    ByRef rateLimitReset As String, _
                                    ByRef responseBody As String, _
                                    ByRef errCode As String, _
                                    ByRef errMsg As String) As Boolean
    On Error GoTo NetErr
#If Mac Then
    ' --- macOS: use curl (no MSXML available) ---
    Dim tmpDir As String, reqFile As String, respFile As String, hdrFile As String
    tmpDir = MacTempDir()
    reqFile = tmpDir & "ghreq.json"
    respFile = tmpDir & "ghresp.json"
    hdrFile = tmpDir & "ghhdr.txt"
    
    WriteTextFile reqFile, jsonBody
    
    Dim cmd As String
    cmd = "curl -s -D " & MacQuote(hdrFile) & " -o " & MacQuote(respFile) & _
          " -w '%{http_code}'" & _
          " -X POST " & MacQuote(GRAPHQL_URL) & _
          " -H 'Content-Type: application/json'" & _
          " -H 'Accept: application/json'" & _
          " -H 'User-Agent: " & DEFAULT_USER_AGENT & "'" & _
          " -H 'Authorization: Bearer " & m_Token & "'" & _
          " --data-binary @" & MacQuote(reqFile)
    
    Dim codeStr As String
    codeStr = MacRunShell(cmd)
    httpStatus = CLng(Val(codeStr))
    responseBody = ReadTextFile(respFile)
    Dim hdrText As String: hdrText = ReadTextFile(hdrFile)
    rateLimitRemaining = ExtractHeader(hdrText, "x-ratelimit-remaining")
    rateLimitReset = ExtractHeader(hdrText, "x-ratelimit-reset")
    SendGraphQLRequest = True
    Exit Function
#Else
    ' --- Windows: MSXML (with test seam) ---
    Dim http As Object
    Set http = CreateHttpClient()                                        ' Req 11.3
    http.setTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, _
                     HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS                    ' Req 9.2
    http.Open "POST", GRAPHQL_URL, False                                ' Req 1.3
    http.setRequestHeader "Content-Type", "application/json"            ' Req 1.4
    http.setRequestHeader "Accept", "application/json"                  ' Req 1.5
    http.setRequestHeader "User-Agent", DEFAULT_USER_AGENT              ' Req 1.6
    http.setRequestHeader "Authorization", "Bearer " & m_Token          ' Req 2.2
    http.send jsonBody
    
    httpStatus = http.Status
    rateLimitRemaining = SafeHeader(http, "X-RateLimit-Remaining")
    rateLimitReset = SafeHeader(http, "X-RateLimit-Reset")
    responseBody = http.responseText
    SendGraphQLRequest = True
    Exit Function
#End If

NetErr:
    ' Distinguish timeout (Err.Number = -2147012894 / 0x80072EE2) from other errors
    If Err.Number = -2147012894 Or InStr(LCase$(Err.Description), "timeout") > 0 Then
        errCode = "TIMEOUT"                                             ' Req 9.3
    Else
        errCode = "NETWORK_ERROR"                                       ' Req 9.1
    End If
    errMsg = Err.Description
    SendGraphQLRequest = False
End Function

#If Mac Then
' macOS helper: run a shell command and return stdout
Private Function MacRunShell(ByVal cmd As String) As String
    On Error Resume Next
    MacRunShell = MacScript("do shell script " & Chr$(34) & _
                            Replace(cmd, Chr$(34), "\" & Chr$(34)) & Chr$(34))
    On Error GoTo 0
End Function

' macOS temp directory (with trailing slash)
Private Function MacTempDir() As String
    Dim d As String
    d = MacRunShell("echo $TMPDIR")
    If Right$(d, 1) <> "/" Then d = d & "/"
    MacTempDir = d
End Function

' Wrap a path in single quotes for shell
Private Function MacQuote(ByVal s As String) As String
    MacQuote = "'" & Replace(s, "'", "'\''") & "'"
End Function

' Write text to file
Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, content
    Close #f
End Sub

' Read text from file
Private Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer: f = FreeFile
    Dim s As String
    On Error Resume Next
    Open path For Input As #f
    s = Input$(LOF(f), f)
    Close #f
    On Error GoTo 0
    ReadTextFile = s
End Function

' Extract a header value (case-insensitive) from raw HTTP headers text
Private Function ExtractHeader(ByVal raw As String, ByVal name As String) As String
    Dim lines() As String, i As Long, ln As String, p As Long
    lines = Split(raw, vbLf)
    For i = 0 To UBound(lines)
        ln = lines(i)
        p = InStr(ln, ":")
        If p > 0 Then
            If LCase$(Trim$(Left$(ln, p - 1))) = LCase$(name) Then
                ExtractHeader = Trim$(Mid$(ln, p + 1))
                Exit Function
            End If
        End If
    Next i
End Function
#End If

' ---------------------------------------------------------------------------
' SetGitHubToken - Configure the Personal Access Token for API authentication
' Accepts an empty string to clear the token.
' Req 2.1
' ---------------------------------------------------------------------------
Public Sub SetGitHubToken(ByVal token As String)
    m_Token = token
End Sub

' ---------------------------------------------------------------------------
' MakeFailureRepoInfo - Create a failure RepoInfo with given error details
' Req 15.1
' ---------------------------------------------------------------------------
Private Function MakeFailureRepoInfo(ByVal errCode As String, ByVal errMsg As String) As RepoInfo
    Dim result As RepoInfo
    result.Success = False
    result.ErrorCode = errCode
    result.ErrorMessage = errMsg
    ' All other fields remain at default (0, "", False)
    MakeFailureRepoInfo = result
End Function

' ---------------------------------------------------------------------------
' MakeFailureIssuesResult - Create a failure RepoIssuesResult with given error details
' Returns a RepoIssuesResult with empty Items array, Success=False, and the
' provided ErrorCode/ErrorMessage.
' Req 17.1-17.21
' ---------------------------------------------------------------------------
Private Sub MakeFailureIssuesResult(ByRef outResult As RepoIssuesResult, ByVal errCode As String, ByVal errMsg As String)
    outResult.Success = False
    outResult.ErrorCode = errCode
    outResult.ErrorMessage = errMsg
    outResult.TotalCount = 0
    m_IssuesCount = 0
End Sub

' ---------------------------------------------------------------------------
' GetIssueItem - Return a specific RepoIssue from the last GetRepoIssues call
' Index is 1-based. Returns a default (empty) RepoIssue if index out of range.
' ---------------------------------------------------------------------------
Public Function GetIssueItem(ByVal index As Long) As RepoIssue
    If index >= 1 And index <= m_IssuesCount Then
        GetIssueItem = m_Issues(index)
    End If
End Function

' ---------------------------------------------------------------------------
' GetRepoInfo - Retrieve repository information from GitHub GraphQL API
' Orchestrates the full flow: validate → build query → send request →
' classify response → parse result.
' Returns a RepoInfo UDT with Success=True on success, or Success=False
' with ErrorCode/ErrorMessage on any failure. Never raises runtime errors.
' Req 1.1, 1.2, 2.3, 2.4, 3.5, 18.1
' ---------------------------------------------------------------------------
Public Function GetRepoInfo(ByVal githubId As String, ByVal repoName As String) As RepoInfo
    On Error GoTo CleanFail
    
    EnsureDefaults
    
    ' Step 1: Check token (Req 2.3, 2.4)
    If m_Token = "" Then
        GetRepoInfo = MakeFailureRepoInfo("MISSING_TOKEN", "GitHub token has not been set. Call SetGitHubToken first.")
        Exit Function
    End If
    
    ' Step 2: Validate githubId (Req 3.1, 3.3)
    If Not ValidateGithubId(githubId) Then
        GetRepoInfo = MakeFailureRepoInfo("INVALID_GITHUB_ID", "Invalid GitHub ID: " & githubId)
        Exit Function
    End If
    
    ' Step 3: Validate repoName (Req 3.2, 3.4)
    If Not ValidateRepoName(repoName) Then
        GetRepoInfo = MakeFailureRepoInfo("INVALID_REPO_NAME", "Invalid repository name: " & repoName)
        Exit Function
    End If
    
    ' Step 4: Build query and request body (Req 4)
    Dim query As String
    Dim jsonBody As String
    query = BuildRepoInfoQuery()
    jsonBody = BuildRequestBody(query, githubId, repoName)
    
    ' Step 5: Send HTTP request (Req 9)
    Dim httpStatus As Long
    Dim rateLimitRemaining As String
    Dim rateLimitReset As String
    Dim responseBody As String
    Dim errCode As String
    Dim errMsg As String
    
    If Not SendGraphQLRequest(jsonBody, httpStatus, rateLimitRemaining, rateLimitReset, responseBody, errCode, errMsg) Then
        GetRepoInfo = MakeFailureRepoInfo(errCode, errMsg)
        Exit Function
    End If
    
    ' Step 6: Classify HTTP response (Req 7, 8)
    Dim httpResult As String
    httpResult = ClassifyHttpResponse(httpStatus, rateLimitRemaining, rateLimitReset, responseBody, errCode, errMsg)
    If httpResult <> "OK" Then
        GetRepoInfo = MakeFailureRepoInfo(errCode, errMsg)
        Exit Function
    End If
    
    ' Step 7: Parse JSON response (Req 10.1)
    Dim parsed As Object
    On Error GoTo ParseErr
    Set parsed = JsonConverter.ParseJson(responseBody)
    On Error GoTo CleanFail
    
    ' Step 8: Classify GraphQL errors (Req 6)
    If ClassifyGraphQLErrors(parsed, githubId, repoName, errCode, errMsg) Then
        GetRepoInfo = MakeFailureRepoInfo(errCode, errMsg)
        Exit Function
    End If
    
    ' Step 9: Parse repository data (Req 5, 10.2, 10.3)
    Dim result As RepoInfo
    Dim parseResult As String
    parseResult = ParseRepository(parsed, githubId, repoName, result)
    If parseResult <> "" Then
        GetRepoInfo = MakeFailureRepoInfo("PARSE_ERROR", "Failed to parse repository data")
        Exit Function
    End If
    
    ' Step 10: Return successful result
    GetRepoInfo = result
    Exit Function

ParseErr:
    GetRepoInfo = MakeFailureRepoInfo("PARSE_ERROR", "Invalid JSON response: " & Err.Description)
    Exit Function

CleanFail:
    GetRepoInfo = MakeFailureRepoInfo("NETWORK_ERROR", Err.Description)
End Function

' ---------------------------------------------------------------------------
' GetRepoIssues - Retrieve repository issues with project fields from GitHub
' Orchestrates the full flow: validate → build query → send request →
' classify response → parse result.
' Populates the ByRef outResult parameter with a RepoIssuesResult UDT.
' Also stores result in m_LastIssuesResult for use by GetRepoIssueField.
' Never raises runtime errors.
' Req 14.1, 14.2, 14.3, 14.4, 14.8, 14.9, 14.10, 14.11, 15.7, 15.8,
'     17.1-17.21
' ---------------------------------------------------------------------------
Public Sub GetRepoIssues(ByVal githubId As String, _
                         ByVal repoName As String, _
                         ByRef outResult As RepoIssuesResult, _
                         Optional ByVal states As String = "ALL")
    On Error GoTo CleanFail
    
    EnsureDefaults
    
    ' Step 1: Check token (Req 17.1, 17.6)
    If m_Token = "" Then
        MakeFailureIssuesResult outResult, "MISSING_TOKEN", _
            "GitHub token has not been set. Call SetGitHubToken first."
        m_LastIssuesResult = outResult
        Exit Sub
    End If
    
    ' Step 2: Validate githubId (Req 17.2, 17.4, 17.6)
    If Not ValidateGithubId(githubId) Then
        MakeFailureIssuesResult outResult, "INVALID_GITHUB_ID", _
            "Invalid GitHub ID: " & githubId
        m_LastIssuesResult = outResult
        Exit Sub
    End If
    
    ' Step 3: Validate repoName (Req 17.3, 17.5, 17.6)
    If Not ValidateRepoName(repoName) Then
        MakeFailureIssuesResult outResult, "INVALID_REPO_NAME", _
            "Invalid repository name: " & repoName
        m_LastIssuesResult = outResult
        Exit Sub
    End If
    
    ' Step 4: Normalize and validate states (Req 14.8, 14.9, 14.10, 14.11)
    states = UCase$(Trim$(states))
    If states <> "OPEN" And states <> "CLOSED" And states <> "ALL" Then
        MakeFailureIssuesResult outResult, "INVALID_STATES", _
            "Invalid states parameter: " & states & ". Must be OPEN, CLOSED, or ALL."
        m_LastIssuesResult = outResult
        Exit Sub
    End If
    
    ' Pagination loop: fetch up to m_MaxIssues issues (100 per page)
    Dim query As String
    query = BuildRepoIssuesQuery(states)
    
    Dim remaining As Long: remaining = m_MaxIssues
    Dim cursor As String: cursor = ""
    Dim hasNextPage As Boolean: hasNextPage = True
    Dim totalFetched As Long: totalFetched = 0
    
    ReDim m_Issues(1 To m_MaxIssues)
    m_IssuesCount = 0
    
    Dim httpStatus As Long
    Dim rateLimitRemaining As String
    Dim rateLimitReset As String
    Dim responseBody As String
    Dim errCode As String
    Dim errMsg As String
    
    Do While remaining > 0 And hasNextPage
        Dim pageSize As Long
        If remaining > 100 Then pageSize = 100 Else pageSize = remaining
        
        ' Build request body with cursor (Req 15.7, 15.8)
        Dim jsonBody As String
        jsonBody = BuildRequestBody(query, githubId, repoName, pageSize, cursor)
        
        ' Send HTTP request (Req 14.3, 14.4, 17.17, 17.18, 17.19)
        If Not SendGraphQLRequest(jsonBody, httpStatus, rateLimitRemaining, _
                                  rateLimitReset, responseBody, errCode, errMsg) Then
            MakeFailureIssuesResult outResult, errCode, errMsg
            m_LastIssuesResult = outResult
            Exit Sub
        End If
        
        ' Classify HTTP response (Req 17.12-17.16)
        Dim httpResult As String
        httpResult = ClassifyHttpResponse(httpStatus, rateLimitRemaining, _
                                          rateLimitReset, responseBody, errCode, errMsg)
        If httpResult <> "OK" Then
            MakeFailureIssuesResult outResult, errCode, errMsg
            m_LastIssuesResult = outResult
            Exit Sub
        End If
        
        ' Parse JSON response (Req 17.20)
        Dim parsed As Object
        On Error GoTo ParseErr
        Set parsed = JsonConverter.ParseJson(responseBody)
        On Error GoTo CleanFail
        
        ' Classify GraphQL errors (Req 17.7-17.11)
        If ClassifyGraphQLErrors(parsed, githubId, repoName, errCode, errMsg) Then
            MakeFailureIssuesResult outResult, errCode, errMsg
            m_LastIssuesResult = outResult
            Exit Sub
        End If
        
        ' Parse issues nodes for this page (Req 16.1-16.19, 17.21)
        Dim issues() As RepoIssue
        Dim issuesCount As Long
        Dim parseResult As String
        parseResult = ParseIssuesNodes(parsed, issues, issuesCount)
        If parseResult <> "" Then
            MakeFailureIssuesResult outResult, "PARSE_ERROR", _
                "Failed to parse issues data"
            m_LastIssuesResult = outResult
            Exit Sub
        End If
        
        ' Append this page's issues into m_Issues
        If issuesCount > 0 Then
            If totalFetched + issuesCount > UBound(m_Issues) Then
                ReDim Preserve m_Issues(1 To totalFetched + issuesCount)
            End If
            Dim pi As Long
            For pi = 1 To issuesCount
                m_Issues(totalFetched + pi) = issues(pi)
            Next pi
            totalFetched = totalFetched + issuesCount
            remaining = remaining - issuesCount
        Else
            Exit Do
        End If
        
        ' Read pageInfo for next iteration
        hasNextPage = False
        cursor = ""
        On Error Resume Next
        hasNextPage = parsed("data")("repository")("issues")("pageInfo")("hasNextPage")
        cursor = parsed("data")("repository")("issues")("pageInfo")("endCursor")
        If Err.Number <> 0 Then hasNextPage = False: cursor = "": Err.Clear
        On Error GoTo CleanFail
    Loop
    
    ' Return successful result
    m_IssuesCount = totalFetched
    outResult.TotalCount = totalFetched
    outResult.Success = True
    outResult.ErrorCode = ""
    outResult.ErrorMessage = ""
    m_LastIssuesResult = outResult
    Exit Sub

ParseErr:
    MakeFailureIssuesResult outResult, "PARSE_ERROR", _
        "Invalid JSON response: " & Err.Description
    m_LastIssuesResult = outResult
    Exit Sub

CleanFail:
    MakeFailureIssuesResult outResult, "NETWORK_ERROR", Err.Description
    m_LastIssuesResult = outResult
End Sub

' ---------------------------------------------------------------------------
' GetRepoField - Excel UDF wrapper to return a single field from RepoInfo
' Returns the value of the specified field, CVErr(xlErrNA) if repo lookup
' fails, or CVErr(xlErrValue) if fieldName is invalid or an error occurs.
' Req 12.1, 12.2, 12.3, 12.4
' ---------------------------------------------------------------------------
Public Function GetRepoField(ByVal githubId As String, _
                             ByVal repoName As String, _
                             ByVal fieldName As String) As Variant
    On Error GoTo UdfFail
    Dim info As RepoInfo
    info = GetRepoInfo(githubId, repoName)

    If Not info.Success Then
        GetRepoField = CVErr(xlErrNA)        ' Req 12.4
        Exit Function
    End If

    Select Case fieldName
        Case "NameWithOwner":   GetRepoField = info.NameWithOwner
        Case "Description":     GetRepoField = info.Description
        Case "PrimaryLanguage": GetRepoField = info.PrimaryLanguage
        Case "Stars":           GetRepoField = info.Stars
        Case "Forks":           GetRepoField = info.Forks
        Case "OpenIssues":      GetRepoField = info.OpenIssues
        Case "DefaultBranch":   GetRepoField = info.DefaultBranch
        Case "Url":             GetRepoField = info.Url
        Case "CreatedAt":       GetRepoField = info.CreatedAt
        Case "UpdatedAt":       GetRepoField = info.UpdatedAt
        Case "IsPrivate":       GetRepoField = info.IsPrivate
        Case "IsFork":          GetRepoField = info.IsFork
        Case "IsArchived":      GetRepoField = info.IsArchived
        Case "DiskUsageKB":     GetRepoField = info.DiskUsageKB
        Case "ProjectsCount":   GetRepoField = info.ProjectsCount
        Case Else:              GetRepoField = CVErr(xlErrValue)   ' Req 12.3
    End Select
    Exit Function

UdfFail:
    GetRepoField = CVErr(xlErrValue)
End Function

' ---------------------------------------------------------------------------
' GetRepoIssueField - Excel UDF wrapper to return a single field from a
' specific issue in the RepoIssuesResult array.
' Parameters:
'   githubId   - GitHub username or organization
'   repoName   - Repository name
'   issueIndex - 1-based index into the Items array
'   fieldName  - Case-insensitive field name (Number, Title, Url, State,
'                Author, CreatedAt, UpdatedAt, ClosedAt, Labels, Assignees,
'                Milestone, ProjectFields)
' Returns:
'   Field value on success, CVErr(xlErrNA) if lookup fails or index is out
'   of range, CVErr(xlErrValue) if fieldName is invalid or an error occurs.
' Req 18.1, 18.2, 18.3, 18.4, 18.5, 18.6
' ---------------------------------------------------------------------------
Public Function GetRepoIssueField(ByVal githubId As String, _
                                  ByVal repoName As String, _
                                  ByVal issueIndex As Long, _
                                  ByVal fieldName As String) As Variant
    On Error GoTo UdfFail
    
    Dim result As RepoIssuesResult
    GetRepoIssues githubId, repoName, result
    
    ' Req 18.5: If GetRepoIssues failed, return #N/A
    If Not result.Success Then
        GetRepoIssueField = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' Req 18.6: If issueIndex out of range, return #N/A
    If issueIndex < 1 Or issueIndex > result.TotalCount Then
        GetRepoIssueField = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' Req 18.3, 18.4: Match fieldName case-insensitively
    Dim issue As RepoIssue
    issue = GetIssueItem(issueIndex)
    
    Select Case UCase$(fieldName)
        Case "NUMBER":        GetRepoIssueField = issue.Number
        Case "TITLE":         GetRepoIssueField = issue.Title
        Case "URL":           GetRepoIssueField = issue.Url
        Case "STATE":         GetRepoIssueField = issue.State
        Case "AUTHOR":        GetRepoIssueField = issue.Author
        Case "CREATEDAT":     GetRepoIssueField = issue.CreatedAt
        Case "UPDATEDAT":     GetRepoIssueField = issue.UpdatedAt
        Case "CLOSEDAT":      GetRepoIssueField = issue.ClosedAt
        Case "LABELS":        GetRepoIssueField = issue.Labels
        Case "ASSIGNEES":     GetRepoIssueField = issue.Assignees
        Case "MILESTONE":     GetRepoIssueField = issue.Milestone
        Case "PROJECTFIELDS": GetRepoIssueField = issue.ProjectFields
        Case Else:            GetRepoIssueField = CVErr(xlErrValue)
    End Select
    Exit Function

UdfFail:
    GetRepoIssueField = CVErr(xlErrValue)
End Function

' ---------------------------------------------------------------------------
' GetRepoIssuesJson - Lấy issues và trả về dưới dạng chuỗi JSON
' Định dạng:
'   {"success":true,"totalCount":N,"errorCode":"","errorMessage":"",
'    "issues":[ {issue}, ... ]}
' Mỗi issue gồm: number, title, url, state, author, createdAt, updatedAt,
'   closedAt, labels, assignees, milestone, projectFields (JSON array thật).
' ---------------------------------------------------------------------------
Public Function GetRepoIssuesJson(ByVal githubId As String, _
                                  ByVal repoName As String, _
                                  Optional ByVal states As String = "ALL") As String
    Dim result As RepoIssuesResult
    GetRepoIssues githubId, repoName, result, states
    
    Dim sb As String
    sb = "{"
    sb = sb & """success"":" & LCase$(CStr(result.Success))
    sb = sb & ",""totalCount"":" & result.TotalCount
    sb = sb & ",""errorCode"":""" & JsonEscape(result.ErrorCode) & """"
    sb = sb & ",""errorMessage"":""" & JsonEscape(result.ErrorMessage) & """"
    sb = sb & ",""issues"":["
    
    If result.Success And result.TotalCount > 0 Then
        Dim i As Long, cur As RepoIssue
        For i = 1 To result.TotalCount
            cur = m_Issues(i)
            If i > 1 Then sb = sb & ","
            sb = sb & "{"
            sb = sb & """number"":" & cur.Number
            sb = sb & ",""title"":""" & JsonEscape(cur.Title) & """"
            sb = sb & ",""url"":""" & JsonEscape(cur.Url) & """"
            sb = sb & ",""state"":""" & JsonEscape(cur.State) & """"
            sb = sb & ",""author"":""" & JsonEscape(cur.Author) & """"
            sb = sb & ",""createdAt"":""" & JsonEscape(FmtDateJson(cur.CreatedAt)) & """"
            sb = sb & ",""updatedAt"":""" & JsonEscape(FmtDateJson(cur.UpdatedAt)) & """"
            sb = sb & ",""closedAt"":""" & JsonEscape(IIf(cur.ClosedAt = CDate(0), "", FmtDateJson(cur.ClosedAt))) & """"
            sb = sb & ",""labels"":""" & JsonEscape(cur.Labels) & """"
            sb = sb & ",""assignees"":""" & JsonEscape(cur.Assignees) & """"
            sb = sb & ",""milestone"":""" & JsonEscape(cur.Milestone) & """"
            ' ProjectFields đã là chuỗi JSON array hợp lệ -> nhúng trực tiếp
            Dim pf As String: pf = cur.ProjectFields
            If pf = "" Then pf = "[]"
            sb = sb & ",""projectFields"":" & pf
            sb = sb & "}"
        Next i
    End If
    
    sb = sb & "]}"
    GetRepoIssuesJson = sb
End Function

' FmtDateJson - Format VBA Date thành chuỗi ISO "yyyy-mm-ddThh:mm:ss" cho JSON
Private Function FmtDateJson(ByVal d As Date) As String
    If d = CDate(0) Then FmtDateJson = "": Exit Function
    FmtDateJson = Format$(d, "yyyy-mm-dd") & "T" & Format$(d, "hh:mm:ss")
End Function

' ---------------------------------------------------------------------------
' WriteRepoIssuesTable - Write issues data to a dedicated worksheet as a table
' Calls GetRepoIssues, creates/replaces a sheet named "{githubId}_{repoName}",
' writes header + data rows starting at A1, and creates a ListObject.
' Raises errors on failure (Sub, not Function).
'
' Parameters:
'   githubId    - GitHub username or organization
'   repoName    - Repository name
'   states      - Optional issue state filter: "ALL" (default), "OPEN", "CLOSED"
'   columns     - Optional array of column names to include (issue fields
'                 and/or project fields in "[ProjectTitle] FieldName" format)
'
' Errors raised:
'   vbObjectError + 1001 - GetRepoIssues returned Success=False
'
' Req 19.1, 19.2, 19.3, 19.4, 19.5, 19.6, 19.7, 19.8, 19.9, 19.10,
'     19.11, 19.12, 19.13, 19.14, 19.15, 19.16, 19.17
' ---------------------------------------------------------------------------
Public Sub WriteRepoIssuesTable(ByVal githubId As String, _
                                ByVal repoName As String, _
                                Optional ByVal states As String = "ALL", _
                                Optional ByVal columns As Variant)
    ' Call GetRepoIssues
    Dim result As RepoIssuesResult
    GetRepoIssues githubId, repoName, result, states
    
    ' Req 19.15: If Success=False, raise error
    If Not result.Success Then
        Err.Raise vbObjectError + ERR_WRITE_FAIL, "WriteRepoIssuesTable", _
                  result.ErrorCode & ": " & result.ErrorMessage
        Exit Sub
    End If
    
    ' Resolve columns
    Dim colNames() As String
    Dim colCount As Long
    colCount = 0
    
    If IsMissing(columns) Or IsEmpty(columns) Then
        ' Default columns: 11 issue fields + discovered project fields (alpha sorted)
        ' Req 19.3, 19.14
        Dim defaultIssueFields(1 To 11) As String
        defaultIssueFields(1) = "Number"
        defaultIssueFields(2) = "Title"
        defaultIssueFields(3) = "Url"
        defaultIssueFields(4) = "State"
        defaultIssueFields(5) = "Author"
        defaultIssueFields(6) = "CreatedAt"
        defaultIssueFields(7) = "UpdatedAt"
        defaultIssueFields(8) = "ClosedAt"
        defaultIssueFields(9) = "Labels"
        defaultIssueFields(10) = "Assignees"
        defaultIssueFields(11) = "Milestone"
        
        ' Discover project fields from data
        Dim projFields() As String
        Dim projFieldCount As Long
        projFieldCount = 0
        
        If result.TotalCount > 0 Then
            ' Iterate all items' ProjectFields JSON to find unique project field names
            ' Use Collection (built-in) for Windows + Mac compatibility
            Dim pfi As Long
            Dim pfParsed As Object
            Dim pfItem As Object
            Dim pfKey As String
            Dim pfCol As New Collection
            
            Dim tmpIssue As RepoIssue
            For pfi = 1 To m_IssuesCount
                tmpIssue = m_Issues(pfi)
                If tmpIssue.ProjectFields <> "[]" And _
                   tmpIssue.ProjectFields <> "" Then
                    On Error Resume Next
                    Set pfParsed = JsonConverter.ParseJson(tmpIssue.ProjectFields)
                    On Error GoTo 0
                    If Not pfParsed Is Nothing Then
                        Dim pfj As Long
                        For pfj = 1 To pfParsed.Count
                            Set pfItem = pfParsed(pfj)
                            pfKey = "[" & pfItem("project") & "] " & pfItem("field")
                            ' Add only if key not already present
                            On Error Resume Next
                            pfCol.Add pfKey, pfKey
                            On Error GoTo 0
                        Next pfj
                    End If
                End If
            Next pfi
            
            ' Sort discovered project fields alphabetically
            If pfCol.Count > 0 Then
                projFieldCount = pfCol.Count
                ReDim projFields(1 To projFieldCount)
                Dim pfIdx As Long
                For pfIdx = 1 To projFieldCount
                    projFields(pfIdx) = pfCol(pfIdx)
                Next pfIdx
                
                ' Simple bubble sort for alphabetical ordering
                Dim sortI As Long, sortJ As Long
                Dim sortTemp As String
                For sortI = 1 To projFieldCount - 1
                    For sortJ = 1 To projFieldCount - sortI
                        If projFields(sortJ) > projFields(sortJ + 1) Then
                            sortTemp = projFields(sortJ)
                            projFields(sortJ) = projFields(sortJ + 1)
                            projFields(sortJ + 1) = sortTemp
                        End If
                    Next sortJ
                Next sortI
            End If
        End If
        
        ' Combine: 11 issue fields + project fields
        colCount = 11 + projFieldCount
        ReDim colNames(1 To colCount)
        Dim ci As Long
        For ci = 1 To 11
            colNames(ci) = defaultIssueFields(ci)
        Next ci
        If projFieldCount > 0 Then
            For ci = 1 To projFieldCount
                colNames(11 + ci) = projFields(ci)
            Next ci
        End If
    Else
        ' User-provided columns array - validate each name
        ' Req 19.4, 19.5, 19.6, 19.7
        Dim userCols As Variant
        userCols = columns
        
        Dim validIssueFields(1 To 11) As String
        validIssueFields(1) = "NUMBER"
        validIssueFields(2) = "TITLE"
        validIssueFields(3) = "URL"
        validIssueFields(4) = "STATE"
        validIssueFields(5) = "AUTHOR"
        validIssueFields(6) = "CREATEDAT"
        validIssueFields(7) = "UPDATEDAT"
        validIssueFields(8) = "CLOSEDAT"
        validIssueFields(9) = "LABELS"
        validIssueFields(10) = "ASSIGNEES"
        validIssueFields(11) = "MILESTONE"
        
        Dim displayIssueFields(1 To 11) As String
        displayIssueFields(1) = "Number"
        displayIssueFields(2) = "Title"
        displayIssueFields(3) = "Url"
        displayIssueFields(4) = "State"
        displayIssueFields(5) = "Author"
        displayIssueFields(6) = "CreatedAt"
        displayIssueFields(7) = "UpdatedAt"
        displayIssueFields(8) = "ClosedAt"
        displayIssueFields(9) = "Labels"
        displayIssueFields(10) = "Assignees"
        displayIssueFields(11) = "Milestone"
        
        ' First pass: count valid columns
        Dim tempCols() As String
        Dim tempCount As Long
        tempCount = 0
        ReDim tempCols(1 To UBound(userCols) - LBound(userCols) + 1)
        
        Dim ui As Long
        For ui = LBound(userCols) To UBound(userCols)
            Dim colName As String
            colName = CStr(userCols(ui))
            
            ' Check if it's a valid issue field (case-insensitive)
            Dim isIssueField As Boolean
            isIssueField = False
            Dim fi As Long
            For fi = 1 To 11
                If UCase$(colName) = validIssueFields(fi) Then
                    isIssueField = True
                    tempCount = tempCount + 1
                    tempCols(tempCount) = displayIssueFields(fi)
                    Exit For
                End If
            Next fi
            
            If Not isIssueField Then
                ' Check if it matches project field format "[...] ..."
                If Left$(colName, 1) = "[" Then
                    Dim closeBracket As Long
                    closeBracket = InStr(2, colName, "]")
                    If closeBracket > 2 And closeBracket < Len(colName) Then
                        ' Valid project field format
                        tempCount = tempCount + 1
                        tempCols(tempCount) = colName
                    End If
                    ' Else: invalid format, skip (Req 19.7)
                End If
                ' Else: not an issue field and doesn't start with [, skip (Req 19.7)
            End If
        Next ui
        
        colCount = tempCount
        If colCount > 0 Then
            ReDim colNames(1 To colCount)
            For ci = 1 To colCount
                colNames(ci) = tempCols(ci)
            Next ci
        End If
    End If
    
    ' If no valid columns found, nothing to write
    If colCount = 0 Then Exit Sub
    
    ' Create/replace a dedicated sheet named after the repo (safe — never touches your data)
    Dim sheetName As String
    sheetName = SafeSheetName(githubId & "_" & repoName)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(sheetName)
    ws.Cells.Clear                          ' Fresh sheet content
    
    Dim startRow As Long, startCol As Long
    startRow = 1
    startCol = 1
    
    Dim tblName As String
    tblName = SafeListObjectName("tblRepoIssues_" & githubId & "_" & repoName)
    
    ' Write header row (Req 19.8)
    Dim c As Long
    For c = 1 To colCount
        ws.Cells(startRow, startCol + c - 1).Value = colNames(c)
    Next c
    
    ' Write data rows (Req 19.10)
    Dim dataRowCount As Long
    dataRowCount = result.TotalCount
    
    If dataRowCount > 0 Then
        Dim r As Long
        Dim curIssue As RepoIssue
        
        For r = 1 To dataRowCount
            curIssue = m_Issues(r)
            
            For c = 1 To colCount
                Dim cellValue As Variant
                cellValue = ""
                
                Dim colUpper As String
                colUpper = UCase$(colNames(c))
                
                ' Check if it's an issue field
                Select Case colUpper
                    Case "NUMBER"
                        cellValue = curIssue.Number
                    Case "TITLE"
                        cellValue = curIssue.Title
                    Case "URL"
                        cellValue = curIssue.Url
                    Case "STATE"
                        cellValue = curIssue.State
                    Case "AUTHOR"
                        cellValue = curIssue.Author
                    Case "CREATEDAT"
                        cellValue = curIssue.CreatedAt
                    Case "UPDATEDAT"
                        cellValue = curIssue.UpdatedAt
                    Case "CLOSEDAT"
                        ' Req 19.11: empty string for Empty ClosedAt
                        If curIssue.ClosedAt = CDate(0) Then
                            cellValue = ""
                        Else
                            cellValue = curIssue.ClosedAt
                        End If
                    Case "LABELS"
                        cellValue = curIssue.Labels
                    Case "ASSIGNEES"
                        cellValue = curIssue.Assignees
                    Case "MILESTONE"
                        cellValue = curIssue.Milestone
                    Case Else
                        ' Project field - parse ProjectFields JSON
                        ' Format: "[ProjectTitle] FieldName"
                        cellValue = GetProjectFieldValue( _
                            curIssue.ProjectFields, colNames(c))
                End Select
                
                ws.Cells(startRow + r, startCol + c - 1).Value = cellValue
            Next c
        Next r
    End If
    
    ' Apply date format to date columns (Req 19.11)
    For c = 1 To colCount
        Select Case UCase$(colNames(c))
            Case "CREATEDAT", "UPDATEDAT", "CLOSEDAT"
                If dataRowCount > 0 Then
                    ws.Range(ws.Cells(startRow + 1, startCol + c - 1), _
                             ws.Cells(startRow + dataRowCount, startCol + c - 1)).NumberFormat = _
                             "yyyy-mm-dd hh:mm:ss"
                End If
        End Select
    Next c
    
    ' Create ListObject (Req 19.12)
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(startRow, startCol), _
                             ws.Cells(startRow + dataRowCount, startCol + colCount - 1))
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = tblName
    
    ' AutoFit all columns (Req 19.17)
    dataRange.Columns.AutoFit
    ws.Activate
End Sub

' ---------------------------------------------------------------------------
' SafeSheetName - Sanitize a string for use as an Excel worksheet name
' Max 31 chars, replaces invalid chars : \ / ? * [ ] with _
' ---------------------------------------------------------------------------
Private Function SafeSheetName(ByVal raw As String) As String
    Dim i As Long, ch As String, result As String
    result = ""
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        Select Case ch
            Case ":", "\", "/", "?", "*", "[", "]"
                result = result & "_"
            Case Else
                result = result & ch
        End Select
    Next i
    If Len(result) > 31 Then result = Left$(result, 31)
    If Len(result) = 0 Then result = "Issues"
    SafeSheetName = result
End Function

' ---------------------------------------------------------------------------
' GetOrCreateSheet - Return the worksheet with given name, creating it if
' it does not exist. Existing sheet content is left for caller to clear.
' ---------------------------------------------------------------------------
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add( _
            After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

' ---------------------------------------------------------------------------
' GetProjectFieldValue - Extract a project field value from ProjectFields JSON
' Given the ProjectFields JSON string and a column name in format
' "[ProjectTitle] FieldName", parses the JSON and returns the matching value.
' Returns empty string if not found or on error.
' Used internally by WriteRepoIssuesTable.
' ---------------------------------------------------------------------------
Private Function GetProjectFieldValue(ByVal projectFieldsJson As String, _
                                      ByVal columnName As String) As String
    GetProjectFieldValue = ""
    
    ' Parse column name to extract ProjectTitle and FieldName
    ' Format: "[ProjectTitle] FieldName"
    If Left$(columnName, 1) <> "[" Then Exit Function
    
    Dim closeBracket As Long
    closeBracket = InStr(2, columnName, "]")
    If closeBracket < 3 Then Exit Function
    
    Dim projTitle As String
    Dim fldName As String
    projTitle = Mid$(columnName, 2, closeBracket - 2)
    fldName = Trim$(Mid$(columnName, closeBracket + 1))
    
    ' Remove leading space from field name if present
    If Left$(fldName, 1) = " " Then fldName = Mid$(fldName, 2)
    
    If Len(projTitle) = 0 Or Len(fldName) = 0 Then Exit Function
    
    ' Parse JSON
    If projectFieldsJson = "[]" Or projectFieldsJson = "" Then Exit Function
    
    Dim parsed As Object
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(projectFieldsJson)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    If parsed Is Nothing Then Exit Function
    
    ' Search for matching project+field combination
    ' ProjectTitle: case-sensitive match per spec
    ' FieldName: case-insensitive match per spec (Req 19.6)
    Dim item As Object
    Dim i As Long
    For i = 1 To parsed.Count
        Set item = parsed(i)
        If item("project") = projTitle Then
            If UCase$(item("field")) = UCase$(fldName) Then
                GetProjectFieldValue = item("value")
                Exit Function
            End If
        End If
    Next i
End Function
