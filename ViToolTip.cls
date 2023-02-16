VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_toolTipHwnd As Long
Private m_parenthWnd As Long

Public Property Get hWnd() As Long
    hWnd = m_toolTipHwnd
End Property

Public Sub Hide()
    SendMessage ByVal m_toolTipHwnd, ByVal TTM_POP, ByVal 0, ByVal 0
End Sub

Public Sub Show()
    Debug.Print "Showing!"
    'SetForegroundWindow m_toolTipHwnd

    'SendMessage ByVal m_toolTipHwnd, ByVal TTM_POP, ByVal 0, ByVal 0
    'SendMessage ByVal m_toolTipHwnd, ByVal TTM_POPUP, ByVal 0, ByVal 0

End Sub

Public Sub AttachWindow(ByVal hWnd As Long)
    m_parenthWnd = hWnd
    
    SetOwner m_toolTipHwnd, hWnd
    
    'm_toolTipHwnd = CreateWindowEx(0, _
        "tooltips_class32", "", _
        0, 5, 5, 200, 200, m_parenthWnd, 0, 0, ByVal 0)
End Sub

Private Sub Class_Initialize()

    m_toolTipHwnd = CreateWindowEx(0, _
        "tooltips_class32", "", _
        0, 0, 0, 0, 0, 0, 0, 0, ByVal 0)
        
    SendMessage ByVal m_toolTipHwnd, ByVal TTM_SETDELAYTIME, 1, 500
    'SendMessage ByVal m_toolTipHwnd, ByVal TTM_SETDELAYTIME, 0, 0
    'SendMessage ByVal m_toolTipHwnd, ByVal TTM_SETDELAYTIME, 2, 0

    SetWindowPos m_toolTipHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End Sub

Public Sub SetToolTip(ByVal szToolTipText As String)

Dim udtTool As TOOLINFO

    ' add control
    With udtTool
        .cbSize = Len(udtTool)
        .hWnd = m_parenthWnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        .uId = m_parenthWnd
        .lpszText = szToolTipText
    End With
    
    SendMessage ByVal m_toolTipHwnd, ByVal TTM_ADDTOOLA, ByVal 0, udtTool
End Sub

Private Sub Class_Terminate()
    
    If m_toolTipHwnd <> 0 Then
        DestroyWindow m_toolTipHwnd
    End If
    
End Sub
