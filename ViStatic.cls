VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowExW Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_SETFONT As Long = &H30
Private Const WM_SETTEXT = &HC

Private hfFont As Long

'local variable(s) to hold property value(s)
Private mvarParentHwnd As Long 'local copy
Private mvarName As String 'local copy
Private mvarHwnd As Long 'local copy
Private mvarCaption As String 'local copy
Private mvarFont As GDIFont 'local copy

Private mvarListCount As Long

Public Property Get ListCount()
    ListCount = mvarListCount
End Property

Public Property Get listIndex() As Long

Dim lngIndex As Long

    lngIndex = SendMessage(mvarHwnd, CB_GETCURSEL, 0&, 0&)
    listIndex = lngIndex

End Property

Public Property Let listIndex(ByVal lngIndex As Long)
    SendMessage mvarHwnd, CB_SETCURSEL, lngIndex, 0&
End Property

Public Function AddString(strData As String)
    SendMessage mvarHwnd, CB_ADDSTRING, 0, StrPtr(strData)
    mvarListCount = mvarListCount + 1
End Function

Public Function Move(X As Long, Y As Long, Width As Long, Height As Long)
    MoveWindow mvarHwnd, X, Y, Width, Height, False
End Function

Public Property Set Font(ByRef vData As GDIFont)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Font = Form1
    Set mvarFont = vData
    SelectFont
End Property

Public Property Get Font() As GDIFont
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Font
    Set Font = mvarFont
End Property

Public Property Let Text(newText As String)

Dim bString() As Byte

    bString = newText
    SendMessage mvarHwnd, WM_SETTEXT, 0&, StrPtr(bString)
    
End Property

Public Property Get hWnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hwnd
    hWnd = mvarHwnd
End Property

Private Function SelectFont()

On Error GoTo Handler

    SelectFont = True
    SendMessage mvarHwnd, WM_SETFONT, mvarFont.Handle, 0
    
    Exit Function
Handler:
    SelectFont = False

End Function

Public Sub Constructor(WindowClass As String, hWnd As Long, Optional AdditionalStyle As Long)

    Class_Terminate
    mvarHwnd = CreateWindowExW(0, StrPtr(WindowClass), StrPtr("Generic_ViEdit"), _
                WS_VISIBLE Or WS_CHILD Or AdditionalStyle, _
                10, 10, 500, 100, hWnd, 0, GetModuleHandle(0), 0)

    If Not g_DefaultFont Is Nothing Then
        Set mvarFont = g_DefaultFont
        SelectFont
    End If

End Sub

Private Sub Class_Terminate()

    If mvarHwnd <> 0 Then
        DestroyWindow mvarHwnd
    End If

End Sub
