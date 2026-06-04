VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' MockHttpClient.cls - Mock MSXML2.ServerXMLHTTP.6.0 for testing
' Mimics the interface used by SendGraphQLRequest in GitHub_API_Client.bas
' Supports test tasks: 1.3, 1.4, 1.5, 1.6, 2.2, 9.1, 9.2, 9.3, 14.3, 14.4
' =============================================================================

' ---------------------------------------------------------------------------
' Tracking fields - record what the caller passed in
' ---------------------------------------------------------------------------
Public LastUrl As String
Public LastMethod As String
Public LastTimeout As Long
Public OpenCallCount As Long
Public LastBody As String

' Headers tracking: stores request headers as "Name: Value" entries
Private m_RequestHeaders As Object  ' Scripting.Dictionary

' ---------------------------------------------------------------------------
' Mock configuration - set these before calling the method under test
' ---------------------------------------------------------------------------
Public MockStatus As Long           ' HTTP status to return from Status property
Public MockBody As String           ' Response body to return from responseText
Public RaiseOnSend As Long          ' If non-zero, send() will Err.Raise this number

' Mock response headers: Dictionary of header name -> value
Private m_ResponseHeaders As Object ' Scripting.Dictionary

' ---------------------------------------------------------------------------
' Class_Initialize
' ---------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_RequestHeaders = CreateObject("Scripting.Dictionary")
    m_RequestHeaders.CompareMode = vbTextCompare  ' Case-insensitive keys
    
    Set m_ResponseHeaders = CreateObject("Scripting.Dictionary")
    m_ResponseHeaders.CompareMode = vbTextCompare ' Case-insensitive keys
    
    MockStatus = 200
    MockBody = ""
    RaiseOnSend = 0
    LastUrl = ""
    LastMethod = ""
    LastTimeout = 0
    OpenCallCount = 0
    LastBody = ""
End Sub

' ---------------------------------------------------------------------------
' Property: LastHeaders - expose request headers dictionary for inspection
' ---------------------------------------------------------------------------
Public Property Get LastHeaders() As Object
    Set LastHeaders = m_RequestHeaders
End Property

' ---------------------------------------------------------------------------
' Property: MockHeaders - expose response headers dictionary for configuration
' ---------------------------------------------------------------------------
Public Property Get MockHeaders() As Object
    Set MockHeaders = m_ResponseHeaders
End Property

' ---------------------------------------------------------------------------
' MSXML2 Interface Methods
' ---------------------------------------------------------------------------

' setTimeouts(resolveTimeout, connectTimeout, sendTimeout, receiveTimeout)
' Stores the last timeout value (sendTimeout) for assertion
Public Sub setTimeouts(ByVal resolveTimeout As Long, _
                       ByVal connectTimeout As Long, _
                       ByVal sendTimeout As Long, _
                       ByVal receiveTimeout As Long)
    LastTimeout = sendTimeout
End Sub

' Open(method, url, async)
' Stores method and url, increments OpenCallCount
Public Sub Open(ByVal method As String, _
                ByVal url As String, _
                ByVal async As Boolean)
    LastMethod = method
    LastUrl = url
    OpenCallCount = OpenCallCount + 1
End Sub

' setRequestHeader(name, value)
' Stores header in the request headers dictionary
Public Sub setRequestHeader(ByVal Name As String, ByVal value As String)
    If m_RequestHeaders.Exists(Name) Then
        m_RequestHeaders(Name) = value
    Else
        m_RequestHeaders.Add Name, value
    End If
End Sub

' send(body)
' If RaiseOnSend <> 0, raises error with that number
' Otherwise stores body in LastBody
Public Sub send(ByVal body As String)
    If RaiseOnSend <> 0 Then
        Err.Raise RaiseOnSend, "MockHttpClient.send", "Simulated send error"
    End If
    LastBody = body
End Sub

' getResponseHeader(name) As String
' Returns value from mock response headers dictionary, or empty string
Public Function getResponseHeader(ByVal Name As String) As String
    If m_ResponseHeaders.Exists(Name) Then
        getResponseHeader = m_ResponseHeaders(Name)
    Else
        getResponseHeader = ""
    End If
End Function

' Status property - returns MockStatus
Public Property Get Status() As Long
    Status = MockStatus
End Property

' responseText property - returns MockBody
Public Property Get responseText() As String
    responseText = MockBody
End Property
