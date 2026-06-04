VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockHttpFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' MockHttpFactory.cls - Factory that creates MockHttpClient instances
' Used with SetHttpFactoryForTest to inject mock HTTP into GitHub_API_Client
' Supports test tasks: 1.3, 1.4, 1.5, 1.6, 2.2, 9.1, 9.2, 9.3, 14.3, 14.4
' =============================================================================

' ---------------------------------------------------------------------------
' Fields
' ---------------------------------------------------------------------------

' Pre-configured client to return (if set); otherwise Create() makes a new one
Private m_PreConfiguredClient As MockHttpClient

' Last client created/returned by Create() - for post-call inspection
Private m_LastCreatedClient As MockHttpClient

' ---------------------------------------------------------------------------
' Property: PreConfiguredClient
' Set this before calling the method under test if you want Create() to
' return a specific pre-configured MockHttpClient instance
' ---------------------------------------------------------------------------
Public Property Set PreConfiguredClient(ByVal client As MockHttpClient)
    Set m_PreConfiguredClient = client
End Property

Public Property Get PreConfiguredClient() As MockHttpClient
    Set PreConfiguredClient = m_PreConfiguredClient
End Property

' ---------------------------------------------------------------------------
' Property: LastCreatedClient (read-only)
' Returns the last client instance that Create() returned, for inspection
' ---------------------------------------------------------------------------
Public Property Get LastCreatedClient() As MockHttpClient
    Set LastCreatedClient = m_LastCreatedClient
End Property

' ---------------------------------------------------------------------------
' Create() As Object
' Returns a MockHttpClient instance.
' If PreConfiguredClient is set, returns that instance.
' Otherwise creates a new MockHttpClient.
' ---------------------------------------------------------------------------
Public Function Create() As Object
    If Not m_PreConfiguredClient Is Nothing Then
        Set m_LastCreatedClient = m_PreConfiguredClient
    Else
        Set m_LastCreatedClient = New MockHttpClient
    End If
    Set Create = m_LastCreatedClient
End Function
