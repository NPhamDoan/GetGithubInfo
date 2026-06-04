Attribute VB_Name = "PBT_Helpers"
Option Explicit

' =============================================================================
' PBT_Helpers.bas - Property-Based Testing Helpers
' Generators and runner for property-based tests
' Supports Properties 1-22
' =============================================================================

' --- RunProperty: PBT runner ---
' Loops `iterations` times, calls generator to get sample, calls predicate with
' sample. When predicate returns False, prints counterexample and raises error.
Public Sub RunProperty(ByVal Name As String, _
                       ByVal iterations As Long, _
                       ByVal generator As String, _
                       ByVal predicate As String)
    Dim i As Long
    Dim sample As Variant
    Dim result As Boolean
    
    Randomize
    
    For i = 1 To iterations
        sample = Application.Run(generator)
        result = Application.Run(predicate, sample)
        If Not result Then
            Debug.Print "FAIL " & Name & " counterexample: " & SampleToString(sample)
            Err.Raise vbObjectError + 9001, Name, "Property failed"
            Exit Sub
        End If
    Next i
End Sub

' --- SampleToString: convert sample to printable string ---
Public Function SampleToString(ByVal sample As Variant) As String
    On Error GoTo Fallback
    
    If IsNull(sample) Then
        SampleToString = "Null"
    ElseIf IsEmpty(sample) Then
        SampleToString = "Empty"
    ElseIf IsArray(sample) Then
        SampleToString = ArrayToString(sample)
    ElseIf IsObject(sample) Then
        SampleToString = "[Object]"
    ElseIf VarType(sample) = vbBoolean Then
        SampleToString = CStr(sample)
    ElseIf VarType(sample) = vbDate Then
        SampleToString = Format$(sample, "yyyy-mm-dd hh:mm:ss")
    Else
        SampleToString = CStr(sample)
    End If
    Exit Function
    
Fallback:
    SampleToString = "[Unprintable: VarType=" & VarType(sample) & "]"
End Function


' --- ArrayToString helper ---
Private Function ArrayToString(ByVal arr As Variant) As String
    Dim i As Long
    Dim s As String
    
    On Error GoTo NotArray
    s = "["
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then s = s & ", "
        s = s & SampleToString(arr(i))
    Next i
    s = s & "]"
    ArrayToString = s
    Exit Function
    
NotArray:
    ArrayToString = "[EmptyArray]"
End Function

' =============================================================================
' GENERATORS
' =============================================================================

' --- GenAnyString: random string with various characters ---
' Includes ASCII printable, control chars (U+0000..U+001F), quotes, backslash,
' and Unicode BMP characters
Public Function GenAnyString() As Variant
    Dim maxLen As Long
    Dim length As Long
    Dim i As Long
    Dim s As String
    Dim charCode As Long
    Dim category As Long
    
    maxLen = 50
    length = Int(Rnd * (maxLen + 1)) ' 0 to maxLen
    s = ""
    
    For i = 1 To length
        category = Int(Rnd * 5)
        Select Case category
            Case 0 ' Control characters U+0000..U+001F
                charCode = Int(Rnd * 32)
            Case 1 ' ASCII printable (32-126)
                charCode = 32 + Int(Rnd * 95)
            Case 2 ' Special chars: quote, backslash
                If Rnd > 0.5 Then
                    charCode = 34 ' "
                Else
                    charCode = 92 ' \
                End If
            Case 3 ' Unicode BMP (256-65535)
                charCode = 256 + Int(Rnd * 65280)
            Case 4 ' Extended ASCII (128-255)
                charCode = 128 + Int(Rnd * 128)
        End Select
        s = s & ChrW$(charCode)
    Next i
    
    GenAnyString = s
End Function

' --- GenWhitespaceString: string with only whitespace chars ---
' Returns empty string or string of spaces/tabs/CR/LF
Public Function GenWhitespaceString() As Variant
    Dim length As Long
    Dim i As Long
    Dim s As String
    Dim wsChars As String
    
    wsChars = " " & vbTab & vbCr & vbLf
    length = Int(Rnd * 11) ' 0 to 10
    s = ""
    
    For i = 1 To length
        s = s & Mid$(wsChars, Int(Rnd * 4) + 1, 1)
    Next i
    
    GenWhitespaceString = s
End Function

' --- GenValidGithubId: string matching [A-Za-z0-9-] with length 1-39 ---
Public Function GenValidGithubId() As Variant
    Dim length As Long
    Dim i As Long
    Dim s As String
    Dim charset As String
    
    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"
    length = 1 + Int(Rnd * 39) ' 1 to 39
    s = ""
    
    For i = 1 To length
        s = s & Mid$(charset, Int(Rnd * Len(charset)) + 1, 1)
    Next i
    
    GenValidGithubId = s
End Function

' --- GenValidRepoName: string matching [A-Za-z0-9._-] with length 1-100 ---
Public Function GenValidRepoName() As Variant
    Dim length As Long
    Dim i As Long
    Dim s As String
    Dim charset As String
    
    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-"
    length = 1 + Int(Rnd * 100) ' 1 to 100
    s = ""
    
    For i = 1 To length
        s = s & Mid$(charset, Int(Rnd * Len(charset)) + 1, 1)
    Next i
    
    GenValidRepoName = s
End Function


' --- GenInvalidGithubId: string that violates charset or length rules ---
' Generates strings that are either too long (>39), contain invalid chars,
' or are empty/whitespace-only
Public Function GenInvalidGithubId() As Variant
    Dim strategy As Long
    Dim s As String
    Dim length As Long
    Dim i As Long
    Dim charset As String
    
    strategy = Int(Rnd * 4)
    
    Select Case strategy
        Case 0 ' Too long (40 to 80 chars, valid charset)
            charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"
            length = 40 + Int(Rnd * 41) ' 40 to 80
            s = ""
            For i = 1 To length
                s = s & Mid$(charset, Int(Rnd * Len(charset)) + 1, 1)
            Next i
            
        Case 1 ' Contains invalid characters (special chars)
            length = 1 + Int(Rnd * 20)
            s = ""
            For i = 1 To length
                If Rnd > 0.5 Then
                    ' Valid char
                    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"
                    s = s & Mid$(charset, Int(Rnd * Len(charset)) + 1, 1)
                Else
                    ' Invalid char: space, dot, underscore, special, unicode
                    Dim invalidChars As String
                    invalidChars = " ._%@!#$&*()+=[]{}|;:',<>?/~`"
                    s = s & Mid$(invalidChars, Int(Rnd * Len(invalidChars)) + 1, 1)
                End If
            Next i
            ' Ensure at least one invalid char is present
            If IsValidGithubIdChars(s) Then
                Dim pos As Long
                pos = Int(Rnd * Len(s)) + 1
                Mid$(s, pos, 1) = "."
            End If
            
        Case 2 ' Empty string
            s = ""
            
        Case 3 ' Whitespace only
            s = CStr(GenWhitespaceString())
            If Len(s) = 0 Then s = " "
    End Select
    
    GenInvalidGithubId = s
End Function

' --- GenInvalidRepoName: string that violates charset or length rules ---
Public Function GenInvalidRepoName() As Variant
    Dim strategy As Long
    Dim s As String
    Dim length As Long
    Dim i As Long
    Dim charset As String
    
    strategy = Int(Rnd * 4)
    
    Select Case strategy
        Case 0 ' Too long (101 to 150 chars, valid charset)
            charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-"
            length = 101 + Int(Rnd * 50) ' 101 to 150
            s = ""
            For i = 1 To length
                s = s & Mid$(charset, Int(Rnd * Len(charset)) + 1, 1)
            Next i
            
        Case 1 ' Contains invalid characters
            length = 1 + Int(Rnd * 20)
            s = ""
            For i = 1 To length
                If Rnd > 0.5 Then
                    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-"
                    s = s & Mid$(charset, Int(Rnd * Len(charset)) + 1, 1)
                Else
                    Dim invChars As String
                    invChars = " %@!#$&*()+=[]{}|;:',<>?/~`"
                    s = s & Mid$(invChars, Int(Rnd * Len(invChars)) + 1, 1)
                End If
            Next i
            ' Ensure at least one invalid char is present
            If IsValidRepoNameChars(s) Then
                Dim pos2 As Long
                pos2 = Int(Rnd * Len(s)) + 1
                Mid$(s, pos2, 1) = " "
            End If
            
        Case 2 ' Empty string
            s = ""
            
        Case 3 ' Whitespace only
            s = CStr(GenWhitespaceString())
            If Len(s) = 0 Then s = " "
    End Select
    
    GenInvalidRepoName = s
End Function

' --- Helper: check if all chars are valid GitHub ID chars ---
Private Function IsValidGithubIdChars(ByVal s As String) As Boolean
    Dim i As Long, ch As Long
    If Len(s) = 0 Then
        IsValidGithubIdChars = False
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or _
                (ch >= 65 And ch <= 90) Or _
                (ch >= 97 And ch <= 122) Or _
                ch = 45) Then
            IsValidGithubIdChars = False
            Exit Function
        End If
    Next i
    IsValidGithubIdChars = True
End Function

' --- Helper: check if all chars are valid repo name chars ---
Private Function IsValidRepoNameChars(ByVal s As String) As Boolean
    Dim i As Long, ch As Long
    If Len(s) = 0 Then
        IsValidRepoNameChars = False
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or _
                (ch >= 65 And ch <= 90) Or _
                (ch >= 97 And ch <= 122) Or _
                ch = 45 Or ch = 46 Or ch = 95) Then
            IsValidRepoNameChars = False
            Exit Function
        End If
    Next i
    IsValidRepoNameChars = True
End Function


' --- GenLong: random Long in [min, max] ---
Public Function GenLong(Optional ByVal minVal As Long = 0, _
                        Optional ByVal maxVal As Long = 1000000) As Variant
    GenLong = CLng(minVal + Int(Rnd * (CDbl(maxVal) - CDbl(minVal) + 1)))
End Function

' --- GenIso8601Date: date string like "2023-08-15T07:23:45Z" ---
' Generates dates between 1970-01-01 and 2100-01-01
Public Function GenIso8601Date() As Variant
    Dim baseDate As Date
    Dim maxSeconds As Double
    Dim randSeconds As Double
    Dim d As Date
    Dim s As String
    Dim useOffset As Boolean
    
    baseDate = #1/1/1970#
    maxSeconds = 4102444800# ' seconds from 1970 to 2100
    randSeconds = Rnd * maxSeconds
    
    d = DateAdd("s", Int(randSeconds), baseDate)
    
    ' Format as ISO 8601
    s = Format$(Year(d), "0000") & "-" & _
        Format$(Month(d), "00") & "-" & _
        Format$(Day(d), "00") & "T" & _
        Format$(Hour(d), "00") & ":" & _
        Format$(Minute(d), "00") & ":" & _
        Format$(Second(d), "00")
    
    ' Sometimes add timezone offset, sometimes Z
    useOffset = (Rnd > 0.7)
    If useOffset Then
        Dim offsetHours As Long
        Dim offsetMinutes As Long
        Dim sign As String
        offsetHours = Int(Rnd * 13) ' 0 to 12
        offsetMinutes = IIf(Rnd > 0.5, 0, 30)
        If Rnd > 0.5 Then
            sign = "+"
        Else
            sign = "-"
        End If
        s = s & sign & Format$(offsetHours, "00") & ":" & Format$(offsetMinutes, "00")
    Else
        s = s & "Z"
    End If
    
    GenIso8601Date = s
End Function

' --- GenHttpStatus: Long in [100, 599] ---
Public Function GenHttpStatus() As Variant
    GenHttpStatus = CLng(100 + Int(Rnd * 500))
End Function

' --- GenGraphQLErrorType: one of known types or random string ---
Public Function GenGraphQLErrorType() As Variant
    Dim choice As Long
    choice = Int(Rnd * 5)
    
    Select Case choice
        Case 0
            GenGraphQLErrorType = "NOT_FOUND"
        Case 1
            GenGraphQLErrorType = "FORBIDDEN"
        Case 2
            GenGraphQLErrorType = "RATE_LIMITED"
        Case Else ' 3, 4 - random string
            Dim length As Long
            Dim i As Long
            Dim s As String
            Dim charset As String
            charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_"
            length = 3 + Int(Rnd * 15)
            s = ""
            For i = 1 To length
                s = s & Mid$(charset, Int(Rnd * Len(charset)) + 1, 1)
            Next i
            GenGraphQLErrorType = s
    End Select
End Function


' --- GenRepoInfo: fully populated RepoInfo UDT ---
Public Function GenRepoInfo() As Variant
    Dim info As RepoInfo
    
    ' Generate NameWithOwner as "owner/repo" format
    info.NameWithOwner = CStr(GenValidGithubId()) & "/" & CStr(GenValidRepoName())
    info.Description = CStr(GenAnyString())
    
    ' PrimaryLanguage: sometimes empty
    If Rnd > 0.3 Then
        Dim langs As Variant
        langs = Array("JavaScript", "Python", "Java", "TypeScript", "C#", _
                      "Go", "Rust", "Ruby", "PHP", "Swift", "Kotlin", "C++")
        info.PrimaryLanguage = langs(Int(Rnd * 12))
    Else
        info.PrimaryLanguage = ""
    End If
    
    info.Stars = CLng(Int(Rnd * 1000000))
    info.Forks = CLng(Int(Rnd * 500000))
    info.OpenIssues = CLng(Int(Rnd * 10000))
    
    ' DefaultBranch: sometimes empty
    If Rnd > 0.2 Then
        Dim branches As Variant
        branches = Array("main", "master", "develop", "dev", "release")
        info.DefaultBranch = branches(Int(Rnd * 5))
    Else
        info.DefaultBranch = ""
    End If
    
    ' Dates between 1970 and 2100
    Dim baseDate As Date
    baseDate = #1/1/1970#
    Dim maxSec As Double
    maxSec = 4102444800#
    info.CreatedAt = DateAdd("s", Int(Rnd * maxSec), baseDate)
    info.UpdatedAt = DateAdd("s", Int(Rnd * maxSec), baseDate)
    ' Ensure UpdatedAt >= CreatedAt
    If info.UpdatedAt < info.CreatedAt Then
        Dim tmp As Date
        tmp = info.CreatedAt
        info.CreatedAt = info.UpdatedAt
        info.UpdatedAt = tmp
    End If
    
    info.Url = "https://github.com/" & info.NameWithOwner
    info.IsPrivate = (Rnd > 0.7)
    info.IsFork = (Rnd > 0.8)
    info.IsArchived = (Rnd > 0.9)
    info.DiskUsageKB = CLng(Int(Rnd * 10000000))
    info.ProjectsCount = CLng(Int(Rnd * 50))
    info.Success = True
    info.ErrorCode = ""
    info.ErrorMessage = ""
    
    GenRepoInfo = info
End Function

' --- GenRepoProject: fully populated RepoProject UDT ---
Public Function GenRepoProject() As Variant
    Dim proj As RepoProject
    
    proj.Number = CLng(1 + Int(Rnd * 1000))
    proj.Title = "Project-" & CStr(Int(Rnd * 10000))
    proj.Url = "https://github.com/users/" & CStr(GenValidGithubId()) & "/projects/" & CStr(proj.Number)
    proj.ShortDescription = CStr(GenAnyString())
    proj.Closed = (Rnd > 0.7)
    proj.Public_ = (Rnd > 0.4)
    
    ' Dates between 1970 and 2100
    Dim baseDate As Date
    baseDate = #1/1/1970#
    Dim maxSec As Double
    maxSec = 4102444800#
    proj.CreatedAt = DateAdd("s", Int(Rnd * maxSec), baseDate)
    proj.UpdatedAt = DateAdd("s", Int(Rnd * maxSec), baseDate)
    ' Ensure UpdatedAt >= CreatedAt
    If proj.UpdatedAt < proj.CreatedAt Then
        Dim tmp As Date
        tmp = proj.CreatedAt
        proj.CreatedAt = proj.UpdatedAt
        proj.UpdatedAt = tmp
    End If
    
    proj.ItemCount = CLng(Int(Rnd * 500))
    
    GenRepoProject = proj
End Function

' --- GenRepoProjectArray: array of RepoProject with size 0 to maxLen ---
Public Function GenRepoProjectArray(Optional ByVal maxLen As Long = 20) As Variant
    Dim count As Long
    Dim i As Long
    Dim arr() As RepoProject
    
    count = Int(Rnd * (maxLen + 1)) ' 0 to maxLen
    
    If count = 0 Then
        GenRepoProjectArray = Array()
        Exit Function
    End If
    
    ReDim arr(0 To count - 1)
    For i = 0 To count - 1
        Dim proj As RepoProject
        Dim v As Variant
        v = GenRepoProject()
        proj = v
        arr(i) = proj
    Next i
    
    GenRepoProjectArray = arr
End Function


' --- GenInvalidJson: takes valid JSON and corrupts it ---
' Strategies: truncate, swap chars, remove braces, inject garbage
Public Function GenInvalidJson() As Variant
    Dim validJson As String
    Dim strategy As Long
    Dim result As String
    
    ' Start with a valid JSON string
    validJson = "{""data"":{""repository"":{""nameWithOwner"":""test/repo""," & _
                """description"":""A test"",""url"":""https://github.com/test/repo""," & _
                """stargazerCount"":42,""forkCount"":10,""isPrivate"":false," & _
                """isFork"":false,""isArchived"":false,""diskUsage"":100," & _
                """createdAt"":""2023-01-01T00:00:00Z"",""updatedAt"":""2023-06-15T12:30:00Z""," & _
                """defaultBranchRef"":{""name"":""main""},""primaryLanguage"":{""name"":""Python""}," & _
                """issues"":{""totalCount"":5},""projectsV2"":{""totalCount"":2}}}}"
    
    strategy = Int(Rnd * 6)
    
    Select Case strategy
        Case 0 ' Truncate at random position
            Dim cutPos As Long
            cutPos = 1 + Int(Rnd * (Len(validJson) - 2))
            result = Left$(validJson, cutPos)
            
        Case 1 ' Swap random characters
            result = validJson
            Dim swapCount As Long
            Dim pos1 As Long
            swapCount = 1 + Int(Rnd * 5)
            Dim j As Long
            For j = 1 To swapCount
                pos1 = 1 + Int(Rnd * Len(result))
                Mid$(result, pos1, 1) = ChrW$(32 + Int(Rnd * 95))
            Next j
            
        Case 2 ' Remove all closing braces
            result = Replace(validJson, "}", "")
            
        Case 3 ' Inject garbage at start
            result = "<<<GARBAGE>>>" & validJson
            
        Case 4 ' Just garbage string
            Dim garbageLen As Long
            garbageLen = 5 + Int(Rnd * 50)
            result = ""
            Dim k As Long
            For k = 1 To garbageLen
                result = result & ChrW$(32 + Int(Rnd * 95))
            Next k
            
        Case 5 ' Empty string
            result = ""
    End Select
    
    GenInvalidJson = result
End Function

' =============================================================================
' HTTP MOCK HELPERS
' =============================================================================

' --- InstallHttpMock: set the test HTTP factory ---
Public Sub InstallHttpMock(ByVal mock As Object)
    Application.Run "GitHub_API_Client.SetHttpFactoryForTest", mock
End Sub

' --- UninstallHttpMock: remove the test HTTP factory ---
Public Sub UninstallHttpMock()
    Application.Run "GitHub_API_Client.SetHttpFactoryForTest", Nothing
End Sub


' =============================================================================
' JSON BUILDER HELPERS
' For building GraphQL response JSON from UDTs (supports tasks 7.2 and 7.4)
' =============================================================================

' --- BuildRepoInfoResponseJson: build complete GitHub GraphQL response JSON ---
' from a RepoInfo UDT
Public Function BuildRepoInfoResponseJson(ByRef info As RepoInfo) As String
    Dim json As String
    
    json = "{""data"":{""repository"":{" & _
           """nameWithOwner"":""" & JsonEscapeHelper(info.NameWithOwner) & """," & _
           """description"":" & JsonStringOrNull(info.Description) & "," & _
           """url"":""" & JsonEscapeHelper(info.Url) & """," & _
           """stargazerCount"":" & CStr(info.Stars) & "," & _
           """forkCount"":" & CStr(info.Forks) & "," & _
           """isPrivate"":" & LCase$(CStr(info.IsPrivate)) & "," & _
           """isFork"":" & LCase$(CStr(info.IsFork)) & "," & _
           """isArchived"":" & LCase$(CStr(info.IsArchived)) & "," & _
           """diskUsage"":" & CStr(info.DiskUsageKB) & "," & _
           """createdAt"":""" & FormatDateIso8601(info.CreatedAt) & """," & _
           """updatedAt"":""" & FormatDateIso8601(info.UpdatedAt) & """," & _
           """defaultBranchRef"":" & JsonObjectOrNull("name", info.DefaultBranch) & "," & _
           """primaryLanguage"":" & JsonObjectOrNull("name", info.PrimaryLanguage) & "," & _
           """issues"":{""totalCount"":" & CStr(info.OpenIssues) & "}," & _
           """projectsV2"":{""totalCount"":" & CStr(info.ProjectsCount) & "}" & _
           "}}}"
    
    BuildRepoInfoResponseJson = json
End Function

' --- BuildProjectsResponseJson: build complete GitHub GraphQL response JSON ---
' from a RepoProject array
Public Function BuildProjectsResponseJson(ByRef arr() As RepoProject) As String
    Dim json As String
    Dim i As Long
    Dim count As Long
    Dim nodesJson As String
    
    ' Determine array size
    On Error Resume Next
    count = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then
        count = 0
    End If
    On Error GoTo 0
    
    ' Build nodes array
    nodesJson = "["
    If count > 0 Then
        For i = LBound(arr) To UBound(arr)
            If i > LBound(arr) Then nodesJson = nodesJson & ","
            nodesJson = nodesJson & _
                "{""number"":" & CStr(arr(i).Number) & "," & _
                """title"":""" & JsonEscapeHelper(arr(i).Title) & """," & _
                """url"":""" & JsonEscapeHelper(arr(i).Url) & """," & _
                """shortDescription"":""" & JsonEscapeHelper(arr(i).ShortDescription) & """," & _
                """closed"":" & LCase$(CStr(arr(i).Closed)) & "," & _
                """public"":" & LCase$(CStr(arr(i).Public_)) & "," & _
                """createdAt"":""" & FormatDateIso8601(arr(i).CreatedAt) & """," & _
                """updatedAt"":""" & FormatDateIso8601(arr(i).UpdatedAt) & """," & _
                """items"":{""totalCount"":" & CStr(arr(i).ItemCount) & "}}"
        Next i
    End If
    nodesJson = nodesJson & "]"
    
    json = "{""data"":{""repository"":{""projectsV2"":{" & _
           """totalCount"":" & CStr(count) & "," & _
           """nodes"":" & nodesJson & _
           "}}}}"
    
    BuildProjectsResponseJson = json
End Function

' =============================================================================
' INTERNAL HELPERS for JSON builders
' =============================================================================

' --- JsonEscapeHelper: escape string for JSON value ---
' Handles ", \, control characters, and other special chars
Public Function JsonEscapeHelper(ByVal s As String) As String
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
    
    JsonEscapeHelper = result
End Function

' --- FormatDateIso8601: format VBA Date as ISO 8601 UTC string ---
Public Function FormatDateIso8601(ByVal d As Date) As String
    FormatDateIso8601 = Format$(Year(d), "0000") & "-" & _
                        Format$(Month(d), "00") & "-" & _
                        Format$(Day(d), "00") & "T" & _
                        Format$(Hour(d), "00") & ":" & _
                        Format$(Minute(d), "00") & ":" & _
                        Format$(Second(d), "00") & "Z"
End Function

' --- JsonStringOrNull: return JSON string or null ---
Private Function JsonStringOrNull(ByVal s As String) As String
    ' We always output the string (even empty) as a JSON string
    ' since the design says empty string maps from null but we need
    ' the parser to handle both null and "" correctly
    If Len(s) = 0 Then
        JsonStringOrNull = "null"
    Else
        JsonStringOrNull = """" & JsonEscapeHelper(s) & """"
    End If
End Function

' --- JsonObjectOrNull: return {"key":"value"} or null ---
Private Function JsonObjectOrNull(ByVal key As String, ByVal value As String) As String
    If Len(value) = 0 Then
        JsonObjectOrNull = "null"
    Else
        JsonObjectOrNull = "{""" & key & """:""" & JsonEscapeHelper(value) & """}"
    End If
End Function
