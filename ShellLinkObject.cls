VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellLinkObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_oWScriptShell As Object
Private m_oShortcut As Object

Private Sub Class_Initialize()
    Set m_oWScriptShell = GetGlobalWScriptShellObject()
End Sub

Public Function Resolve(ByVal szPath As String) As Boolean
    On Error GoTo Handler
    
    'If Not FileExists(szPath) Then Exit Function
    Set m_oShortcut = m_oWScriptShell.CreateShortcut(szPath)
    
    'Fix target on Windows 8 machines
    If g_Windows8 Or g_Windows81 Then
        If Not FileExists(Me.Target) Then
            If FileExists(Wow64Wrapper(Me.Target)) Then
                Me.Target = Wow64Wrapper(Me.Target)
                Me.Save
                
                Set m_oShortcut = m_oWScriptShell.CreateShortcut(szPath)
            End If
        End If
    End If
    
    Resolve = True
    Exit Function
Handler:
End Function

Public Function Save()
    If m_oShortcut Is Nothing Then
        Exit Function
    End If
    
    m_oShortcut.Save
End Function

Public Property Get Arguments() As String
    If m_oShortcut Is Nothing Then
        Exit Property
    End If
    
    Arguements = m_oShortcut.Arguments
End Property

Public Property Let Target(ByVal newTarget As String)
    If m_oShortcut Is Nothing Then
        Exit Property
    End If
    
    m_oShortcut.TargetPath = newTarget
End Property

Public Property Get Target() As String
    If m_oShortcut Is Nothing Then
        Exit Property
    End If
    
    Target = m_oShortcut.TargetPath
End Property

Public Function GetIconLocation(ByRef szSBS As String) As Long
    On Error GoTo Handler

Dim szIconLocation() As String

    If m_oShortcut Is Nothing Then
        Exit Function
    End If
    
    If InStr(CStr(m_oShortcut.IconLocation), ",") > 0 Then
        szIconLocation = Split(m_oShortcut.IconLocation, ",")
        
        szSBS = CStr(szIconLocation(0))
        GetIconLocation = CLng(szIconLocation(1))
    Else
        szSBS = m_oShortcut.IconLocation
    End If
    
Handler:
End Function
