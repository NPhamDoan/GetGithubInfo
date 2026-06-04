Attribute VB_Name = "GitHub_API_Client_Excel"
Option Explicit

' ===========================================================================
' GitHub_API_Client_Excel - Compact module for WriteRepoIssuesTable only
' Requires: JsonConverter.bas, Microsoft Scripting Runtime reference
' ===========================================================================

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

Private m_Token As String
Private m_MaxIssues As Long
Private m_Issues() As RepoIssue
Private m_IssuesCount As Long

Private Const HTTP_TIMEOUT_MS As Long = 30000
Private Const GRAPHQL_URL As String = "https://api.github.com/graphql"
Private Const USER_AGENT As String = "VBA-GitHub-GraphQL-Client"
Private Const DEFAULT_MAX_ISSUES As Long = 100

' ===========================================================================
' PUBLIC API
' ===========================================================================

Public Sub SetGitHubToken(ByVal token As String)
    m_Token = token
End Sub

Public Sub SetMaxIssues(ByVal maxIssues As Long)
    If maxIssues < 1 Then
        Err.Raise vbObjectError + 1003, "SetMaxIssues", "maxIssues must be >= 1"
    End If
    m_MaxIssues = maxIssues
End Sub

Private Sub GetRepoIssues(ByVal githubId As String, _
                         ByVal repoName As String, _
                         ByRef outResult As RepoIssuesResult, _
                         Optional ByVal states As String = "ALL")
    On Error GoTo CleanFail
    If m_MaxIssues = 0 Then m_MaxIssues = DEFAULT_MAX_ISSUES
    
    ' Validate inputs
    If m_Token = "" Then
        SetFailure outResult, "MISSING_TOKEN", "Call SetGitHubToken first."
        Exit Sub
    End If
    If Not IsValidGithubId(githubId) Then
        SetFailure outResult, "INVALID_GITHUB_ID", "Invalid GitHub ID: " & githubId
        Exit Sub
    End If
    If Not IsValidRepoName(repoName) Then
        SetFailure outResult, "INVALID_REPO_NAME", "Invalid repo name: " & repoName
        Exit Sub
    End If
    states = UCase$(Trim$(states))
    If states <> "OPEN" And states <> "CLOSED" And states <> "ALL" Then
        SetFailure outResult, "INVALID_STATES", "states must be OPEN, CLOSED, or ALL"
        Exit Sub
    End If
    
    ' Pagination loop
    Dim totalFetched As Long: totalFetched = 0
    Dim remaining As Long: remaining = m_MaxIssues
    Dim cursor As String: cursor = ""
    Dim hasNextPage As Boolean: hasNextPage = True
    
    m_IssuesCount = 0
    ReDim m_Issues(1 To m_MaxIssues)
    
    Do While remaining > 0 And hasNextPage
        Dim pageSize As Long
        If remaining > 100 Then pageSize = 100 Else pageSize = remaining
        
        ' Build request body with cursor
        Dim query As String, body As String
        query = BuildIssuesQuery(states)
        body = "{""query"":""" & JEsc(query) & """,""variables"":{""owner"":""" & JEsc(githubId) & """,""name"":""" & JEsc(repoName) & """,""first"":" & pageSize
        If cursor <> "" Then body = body & ",""after"":""" & JEsc(cursor) & """"
        body = body & "}}"
        
        ' Send HTTP request
        Dim http As Object, httpStatus As Long
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
        http.Open "POST", GRAPHQL_URL, False
        http.setRequestHeader "Content-Type", "application/json"
        http.setRequestHeader "Accept", "application/json"
        http.setRequestHeader "User-Agent", USER_AGENT
        http.setRequestHeader "Authorization", "Bearer " & m_Token
        http.send body
        httpStatus = http.Status
        
        ' Check HTTP status
        If httpStatus <> 200 Then
            Dim rlr As String: rlr = "": On Error Resume Next: rlr = http.getResponseHeader("X-RateLimit-Remaining"): On Error GoTo CleanFail
            Select Case httpStatus
                Case 401: SetFailure outResult, "UNAUTHORIZED", "Token invalid or expired": Exit Sub
                Case 403
                    If rlr = "0" Then SetFailure outResult, "RATE_LIMITED", "Rate limit exceeded" Else SetFailure outResult, "FORBIDDEN", "Access denied"
                    Exit Sub
                Case 500 To 599: SetFailure outResult, "SERVER_ERROR", "HTTP " & httpStatus: Exit Sub
                Case Else: SetFailure outResult, "HTTP_ERROR", "HTTP " & httpStatus: Exit Sub
            End Select
        End If
        
        ' Parse JSON
        Dim parsed As Object
        On Error GoTo ParseErr
        Set parsed = JsonConverter.ParseJson(http.responseText)
        On Error GoTo CleanFail
        
        ' Check GraphQL errors
        If parsed.Exists("errors") Then
            If parsed("errors").Count > 0 Then
                Dim eType As String, eMsg As String
                On Error Resume Next
                eType = parsed("errors")(1)("type"): eMsg = parsed("errors")(1)("message")
                On Error GoTo CleanFail
                Select Case eType
                    Case "NOT_FOUND": SetFailure outResult, "NOT_FOUND", githubId & "/" & repoName
                    Case "FORBIDDEN": SetFailure outResult, "FORBIDDEN", eMsg
                    Case "RATE_LIMITED": SetFailure outResult, "RATE_LIMITED", eMsg
                    Case Else: SetFailure outResult, "GRAPHQL_ERROR", eMsg
                End Select
                Exit Sub
            End If
        End If
        
        ' Check data.repository exists
        Dim repo As Object
        On Error Resume Next
        Set repo = parsed("data")("repository")
        If Err.Number <> 0 Or repo Is Nothing Then
            On Error GoTo CleanFail
            SetFailure outResult, "NOT_FOUND", githubId & "/" & repoName
            Exit Sub
        End If
        On Error GoTo CleanFail
        
        ' Get pageInfo for pagination
        Dim issuesObj As Object
        Set issuesObj = parsed("data")("repository")("issues")
        
        On Error Resume Next
        hasNextPage = issuesObj("pageInfo")("hasNextPage")
        cursor = issuesObj("pageInfo")("endCursor")
        If Err.Number <> 0 Then hasNextPage = False: cursor = "": Err.Clear
        On Error GoTo CleanFail
        
        ' Parse nodes
        Dim nodes As Object, nodeCount As Long
        Set nodes = issuesObj("nodes")
        nodeCount = nodes.Count
        
        If nodeCount = 0 Then Exit Do
        
        ' Expand array if needed
        If totalFetched + nodeCount > UBound(m_Issues) Then
            ReDim Preserve m_Issues(1 To totalFetched + nodeCount + 100)
        End If
        
        ' Parse each issue node
        Dim i As Long, j As Long, node As Object
        For i = 1 To nodeCount
            totalFetched = totalFetched + 1
            Set node = nodes(i)
            m_Issues(totalFetched).Number = CLng(node("number"))
            m_Issues(totalFetched).Title = CStr(node("title"))
            m_Issues(totalFetched).Url = CStr(node("url"))
            m_Issues(totalFetched).State = CStr(node("state"))
            
            On Error Resume Next
            m_Issues(totalFetched).Author = node("author")("login")
            If Err.Number <> 0 Then m_Issues(totalFetched).Author = "": Err.Clear
            On Error GoTo CleanFail
            
            m_Issues(totalFetched).CreatedAt = ParseISO(CStr(node("createdAt")))
            m_Issues(totalFetched).UpdatedAt = ParseISO(CStr(node("updatedAt")))
            Dim ca As Variant
            On Error Resume Next: ca = node("closedAt"): On Error GoTo CleanFail
            If Not IsNull(ca) Then m_Issues(totalFetched).ClosedAt = ParseISO(CStr(ca))
            
            Dim lbls As Object, lStr As String: lStr = ""
            On Error Resume Next: Set lbls = node("labels")("nodes"): On Error GoTo CleanFail
            If Not lbls Is Nothing Then
                For j = 1 To lbls.Count
                    If j > 1 Then lStr = lStr & ", "
                    lStr = lStr & CStr(lbls(j)("name"))
                Next j
            End If
            m_Issues(totalFetched).Labels = lStr
            
            Dim asns As Object, aStr As String: aStr = ""
            On Error Resume Next: Set asns = node("assignees")("nodes"): On Error GoTo CleanFail
            If Not asns Is Nothing Then
                For j = 1 To asns.Count
                    If j > 1 Then aStr = aStr & ", "
                    aStr = aStr & CStr(asns(j)("login"))
                Next j
            End If
            m_Issues(totalFetched).Assignees = aStr
            
            On Error Resume Next
            m_Issues(totalFetched).Milestone = node("milestone")("title")
            If Err.Number <> 0 Then m_Issues(totalFetched).Milestone = "": Err.Clear
            On Error GoTo CleanFail
            
            ' ProjectFields
            Dim piNodes As Object
            On Error Resume Next: Set piNodes = node("projectItems")("nodes"): On Error GoTo CleanFail
            If piNodes Is Nothing Or piNodes.Count = 0 Then
                m_Issues(totalFetched).ProjectFields = "[]"
            Else
                Dim pf As String, pfFirst As Boolean: pf = "[": pfFirst = True
                Dim k As Long, mm As Long, piNode As Object, fvNodes As Object, fvNode As Object
                Dim pTitle As String, fName As String, fVal As String
                For k = 1 To piNodes.Count
                    Set piNode = piNodes(k)
                    On Error Resume Next: pTitle = piNode("project")("title"): If Err.Number <> 0 Then pTitle = "": Err.Clear
                    Set fvNodes = piNode("fieldValues")("nodes"): On Error GoTo CleanFail
                    If Not fvNodes Is Nothing Then
                        For mm = 1 To fvNodes.Count
                            Set fvNode = fvNodes(mm)
                            fName = "": fVal = ""
                            On Error Resume Next
                            If fvNode.Exists("text") Then
                                fName = fvNode("field")("name"): fVal = fvNode("text")
                            ElseIf fvNode.Exists("number") Then
                                fName = fvNode("field")("name"): fVal = CStr(fvNode("number"))
                            ElseIf fvNode.Exists("date") Then
                                fName = fvNode("field")("name"): fVal = fvNode("date")
                            ElseIf fvNode.Exists("title") Then
                                fName = fvNode("field")("name"): fVal = fvNode("title")
                            ElseIf fvNode.Exists("name") And fvNode.Exists("field") Then
                                fName = fvNode("field")("name"): fVal = fvNode("name")
                            End If
                            Err.Clear: On Error GoTo CleanFail
                            If fName <> "" Then
                                If Not pfFirst Then pf = pf & ","
                                pf = pf & "{""project"":""" & JEsc(pTitle) & """,""field"":""" & JEsc(fName) & """,""value"":""" & JEsc(fVal) & """}"
                                pfFirst = False
                            End If
                        Next mm
                    End If
                Next k
                pf = pf & "]"
                m_Issues(totalFetched).ProjectFields = pf
            End If
        Next i
        
        remaining = remaining - nodeCount
    Loop
    
    m_IssuesCount = totalFetched
    outResult.TotalCount = totalFetched
    outResult.Success = True
    outResult.ErrorCode = ""
    outResult.ErrorMessage = ""
    Exit Sub
ParseErr:
    SetFailure outResult, "PARSE_ERROR", "Invalid JSON: " & Err.Description
    Exit Sub
CleanFail:
    SetFailure outResult, "NETWORK_ERROR", Err.Description
End Sub

Public Sub WriteRepoIssuesTable(ByVal githubId As String, _
                                ByVal repoName As String, _
                                Optional ByVal states As String = "ALL", _
                                Optional ByVal columns As Variant)
    Dim result As RepoIssuesResult
    GetRepoIssues githubId, repoName, result, states
    If Not result.Success Then
        Err.Raise vbObjectError + 1001, "WriteRepoIssuesTable", result.ErrorCode & ": " & result.ErrorMessage
    End If
    
    ' Resolve columns
    Dim colNames() As String, colCount As Long
    colCount = 0
    
    If IsMissing(columns) Or IsEmpty(columns) Then
        ' Default: 11 issue fields + discovered project fields
        Dim defCols(1 To 11) As String
        defCols(1) = "Number": defCols(2) = "Title": defCols(3) = "Url"
        defCols(4) = "State": defCols(5) = "Author": defCols(6) = "CreatedAt"
        defCols(7) = "UpdatedAt": defCols(8) = "ClosedAt": defCols(9) = "Labels"
        defCols(10) = "Assignees": defCols(11) = "Milestone"
        
        ' Discover project fields
        Dim pfDict As Object: Set pfDict = CreateObject("Scripting.Dictionary")
        Dim pi As Long, pfParsed As Object, pfItem As Object, pfKey As String
        Dim tmpIssue As RepoIssue
        For pi = 1 To m_IssuesCount
            tmpIssue = m_Issues(pi)
            If tmpIssue.ProjectFields <> "[]" And tmpIssue.ProjectFields <> "" Then
                On Error Resume Next
                Set pfParsed = JsonConverter.ParseJson(tmpIssue.ProjectFields)
                On Error GoTo 0
                If Not pfParsed Is Nothing Then
                    Dim pj As Long
                    For pj = 1 To pfParsed.Count
                        Set pfItem = pfParsed(pj)
                        pfKey = "[" & pfItem("project") & "] " & pfItem("field")
                        If Not pfDict.Exists(pfKey) Then pfDict.Add pfKey, pfKey
                    Next pj
                End If
            End If
        Next pi
        
        colCount = 11 + pfDict.Count
        ReDim colNames(1 To colCount)
        Dim ci As Long
        For ci = 1 To 11: colNames(ci) = defCols(ci): Next ci
        If pfDict.Count > 0 Then
            Dim pfKeys As Variant: pfKeys = pfDict.Keys
            ' Sort alphabetically
            Dim si As Long, sj As Long, st As String
            For si = 0 To UBound(pfKeys) - 1
                For sj = 0 To UBound(pfKeys) - si - 1
                    If pfKeys(sj) > pfKeys(sj + 1) Then
                        st = pfKeys(sj): pfKeys(sj) = pfKeys(sj + 1): pfKeys(sj + 1) = st
                    End If
                Next sj
            Next si
            For ci = 0 To UBound(pfKeys): colNames(12 + ci) = pfKeys(ci): Next ci
        End If
    Else
        ' User-specified columns
        Dim validFields(1 To 11) As String
        validFields(1) = "NUMBER": validFields(2) = "TITLE": validFields(3) = "URL"
        validFields(4) = "STATE": validFields(5) = "AUTHOR": validFields(6) = "CREATEDAT"
        validFields(7) = "UPDATEDAT": validFields(8) = "CLOSEDAT": validFields(9) = "LABELS"
        validFields(10) = "ASSIGNEES": validFields(11) = "MILESTONE"
        Dim dispFields(1 To 11) As String
        dispFields(1) = "Number": dispFields(2) = "Title": dispFields(3) = "Url"
        dispFields(4) = "State": dispFields(5) = "Author": dispFields(6) = "CreatedAt"
        dispFields(7) = "UpdatedAt": dispFields(8) = "ClosedAt": dispFields(9) = "Labels"
        dispFields(10) = "Assignees": dispFields(11) = "Milestone"
        
        Dim tmpCols() As String, tmpCount As Long: tmpCount = 0
        ReDim tmpCols(1 To UBound(columns) - LBound(columns) + 1)
        Dim ui As Long, fi As Long, cName As String, isIF As Boolean
        For ui = LBound(columns) To UBound(columns)
            cName = CStr(columns(ui)): isIF = False
            For fi = 1 To 11
                If UCase$(cName) = validFields(fi) Then
                    isIF = True: tmpCount = tmpCount + 1: tmpCols(tmpCount) = dispFields(fi): Exit For
                End If
            Next fi
            If Not isIF And Left$(cName, 1) = "[" Then
                If InStr(2, cName, "]") > 2 Then tmpCount = tmpCount + 1: tmpCols(tmpCount) = cName
            End If
        Next ui
        colCount = tmpCount
        If colCount > 0 Then ReDim colNames(1 To colCount): For ci = 1 To colCount: colNames(ci) = tmpCols(ci): Next ci
    End If
    
    If colCount = 0 Then Exit Sub
    
    ' Create/replace a dedicated sheet named after the repo (safe — never touches your data)
    Dim sheetName As String: sheetName = SafeSheetName(githubId & "_" & repoName)
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(sheetName)
    ws.Cells.Clear                              ' Fresh sheet content
    
    Dim sRow As Long: sRow = 1
    Dim sCol As Long: sCol = 1
    Dim c As Long
    Dim tblName As String: tblName = SafeName("tblRepoIssues_" & githubId & "_" & repoName)
    
    ' Header row
    For c = 1 To colCount: ws.Cells(sRow, sCol + c - 1).Value = colNames(c): Next c
    
    ' Data rows
    Dim r As Long, cur As RepoIssue, cv As Variant
    For r = 1 To result.TotalCount
        cur = m_Issues(r)
        For c = 1 To colCount
            cv = ""
            Select Case UCase$(colNames(c))
                Case "NUMBER": cv = cur.Number
                Case "TITLE": cv = cur.Title
                Case "URL": cv = cur.Url
                Case "STATE": cv = cur.State
                Case "AUTHOR": cv = cur.Author
                Case "CREATEDAT": cv = cur.CreatedAt
                Case "UPDATEDAT": cv = cur.UpdatedAt
                Case "CLOSEDAT": If cur.ClosedAt = CDate(0) Then cv = "" Else cv = cur.ClosedAt
                Case "LABELS": cv = cur.Labels
                Case "ASSIGNEES": cv = cur.Assignees
                Case "MILESTONE": cv = cur.Milestone
                Case Else: cv = GetPFValue(cur.ProjectFields, colNames(c))
            End Select
            ws.Cells(sRow + r, sCol + c - 1).Value = cv
        Next c
    Next r
    
    ' Format date columns
    For c = 1 To colCount
        Select Case UCase$(colNames(c))
            Case "CREATEDAT", "UPDATEDAT", "CLOSEDAT"
                If result.TotalCount > 0 Then
                    ws.Range(ws.Cells(sRow + 1, sCol + c - 1), ws.Cells(sRow + result.TotalCount, sCol + c - 1)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End If
        End Select
    Next c
    
    ' Create ListObject
    Dim dataRng As Range
    Set dataRng = ws.Range(ws.Cells(sRow, sCol), ws.Cells(sRow + result.TotalCount, sCol + colCount - 1))
    Dim lo As ListObject: Set lo = ws.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
    lo.Name = tblName
    dataRng.Columns.AutoFit
    ws.Activate
End Sub

' ===========================================================================
' PRIVATE HELPERS
' ===========================================================================

Private Sub SetFailure(ByRef r As RepoIssuesResult, ByVal code As String, ByVal msg As String)
    r.Success = False: r.ErrorCode = code: r.ErrorMessage = msg: r.TotalCount = 0: m_IssuesCount = 0
End Sub

' Sanitize a string for use as an Excel sheet name (max 31 chars, no : \ / ? * [ ])
Private Function SafeSheetName(ByVal raw As String) As String
    Dim i As Long, ch As String, r As String: r = ""
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        Select Case ch
            Case ":", "\", "/", "?", "*", "[", "]": r = r & "_"
            Case Else: r = r & ch
        End Select
    Next i
    If Len(r) > 31 Then r = Left$(r, 31)
    If Len(r) = 0 Then r = "Issues"
    SafeSheetName = r
End Function

' Return existing sheet with given name, or create a new one. Existing data is replaced.
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function IsValidGithubId(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Or Len(s) > 39 Then Exit Function
    Dim i As Long, ch As Long
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or ch = 45) Then Exit Function
    Next i
    IsValidGithubId = True
End Function

Private Function IsValidRepoName(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Or Len(s) > 100 Then Exit Function
    Dim i As Long, ch As Long
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or ch = 45 Or ch = 46 Or ch = 95) Then Exit Function
    Next i
    IsValidRepoName = True
End Function

Private Function JEsc(ByVal s As String) As String
    Dim i As Long, ch As Long, r As String: r = ""
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        Select Case ch
            Case 34: r = r & "\""": Case 92: r = r & "\\"
            Case 10: r = r & "\n": Case 13: r = r & "\r": Case 9: r = r & "\t"
            Case 0 To 31: r = r & "\u" & Right$("0000" & Hex$(ch), 4)
            Case Else: r = r & ChrW$(ch)
        End Select
    Next i
    JEsc = r
End Function

Private Function BuildIssuesQuery(ByVal states As String) As String
    Dim sf As String
    Select Case states
        Case "OPEN": sf = "[OPEN]"
        Case "CLOSED": sf = "[CLOSED]"
        Case Else: sf = "[OPEN, CLOSED]"
    End Select
    BuildIssuesQuery = "query($owner:String!,$name:String!,$first:Int!,$after:String){repository(owner:$owner,name:$name){issues(first:$first,after:$after,states:" & sf & ",orderBy:{field:CREATED_AT,direction:DESC}){pageInfo{hasNextPage endCursor}totalCount nodes{number title url state author{login}createdAt updatedAt closedAt labels(first:10){nodes{name}}assignees(first:5){nodes{login}}milestone{title}projectItems(first:10){nodes{project{title}fieldValues(first:20){nodes{...on ProjectV2ItemFieldSingleSelectValue{field{...on ProjectV2SingleSelectField{name}}name}...on ProjectV2ItemFieldTextValue{field{...on ProjectV2Field{name}}text}...on ProjectV2ItemFieldNumberValue{field{...on ProjectV2Field{name}}number}...on ProjectV2ItemFieldDateValue{field{...on ProjectV2Field{name}}date}...on ProjectV2ItemFieldIterationValue{field{...on ProjectV2IterationField{name}}title}}}}}}}}}"
End Function

Private Function ParseISO(ByVal s As String) As Date
    On Error Resume Next
    Dim t As Long: t = InStr(1, s, "T")
    If t = 0 Then ParseISO = 0: Exit Function
    ParseISO = DateSerial(CLng(Mid$(s, 1, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2))) + _
               TimeSerial(CLng(Mid$(s, t + 1, 2)), CLng(Mid$(s, t + 4, 2)), CLng(Mid$(s, t + 7, 2)))
    If Err.Number <> 0 Then ParseISO = 0
    On Error GoTo 0
End Function

Private Function SafeName(ByVal raw As String) As String
    Dim i As Long, ch As Long, r As String: r = ""
    For i = 1 To Len(raw)
        ch = AscW(Mid$(raw, i, 1))
        If (ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or ch = 95 Then
            r = r & ChrW$(ch)
        Else: r = r & "_"
        End If
    Next i
    If Len(r) > 0 Then: If AscW(Left$(r, 1)) >= 48 And AscW(Left$(r, 1)) <= 57 Then r = "_" & r
    If Len(r) > 255 Then r = Left$(r, 255)
    SafeName = r
End Function

Private Function GetPFValue(ByVal pfJson As String, ByVal colName As String) As String
    GetPFValue = ""
    If Left$(colName, 1) <> "[" Then Exit Function
    Dim cb As Long: cb = InStr(2, colName, "]")
    If cb < 3 Then Exit Function
    Dim pt As String: pt = Mid$(colName, 2, cb - 2)
    Dim fn As String: fn = Trim$(Mid$(colName, cb + 1))
    If Len(pt) = 0 Or Len(fn) = 0 Then Exit Function
    If pfJson = "[]" Or pfJson = "" Then Exit Function
    Dim p As Object
    On Error Resume Next: Set p = JsonConverter.ParseJson(pfJson)
    If Err.Number <> 0 Then On Error GoTo 0: Exit Function
    On Error GoTo 0
    Dim item As Object, i As Long
    For i = 1 To p.Count
        Set item = p(i)
        If item("project") = pt And UCase$(item("field")) = UCase$(fn) Then
            GetPFValue = item("value"): Exit Function
        End If
    Next i
End Function
