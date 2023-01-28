VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VistaSearchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowExW Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long

'Dont remove this, it screws up unicode shit
Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, _
   ByVal msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Implements IHookSink

Public Event onFocus()
Public Event onLostFocus()
Public Event onKeyDown(ByRef KeyCode As Long)
Public Event onKeyUp(ByRef KeyCode As Long)

Public Event onChange()
Public Event onMouseWheel(ByVal wParam As Long)

Private m_hEdit As Long
Private m_hDC As Long
Private m_font As GDIFont

Private m_defWindowProc As Long
Private m_hasFocus As Boolean

Private WithEvents m_container As frmZOrderKeeper
Attribute m_container.VB_VarHelpID = -1

Private m_BackColour As Long
Private m_ForeColour As Long
Private m_FocusColour As Long

'Windows XP Box-Fix
Private m_WindowWidth As Long
Private m_WindowHeight As Long
Private m_backgroundBrush As GDIBrush

Public Property Let FocusColour(newFocusColour As Long)
    m_FocusColour = newFocusColour
End Property

Public Property Let ForeColour(newForeColour As Long)
    m_ForeColour = newForeColour
    If m_hEdit <> 0 Then SetTextColor m_hEdit, m_ForeColour
End Property

Public Property Let BackColour(newBackColour As Long)
    m_BackColour = newBackColour

    If m_hEdit <> 0 Then SetBkColor m_hEdit, m_BackColour
    If Not m_backgroundBrush Is Nothing Then m_backgroundBrush.Colour = newBackColour
End Property

Public Function SetKeyboardFocus()
    win.SetFocus m_hEdit
End Function

Public Property Get Font() As GDIFont
    Set Font = m_font
End Property

Public Property Let Font(ByRef newFont As GDIFont)
    If newFont Is Nothing Then Exit Property

    Set m_font = newFont

    SendMessageW m_hEdit, WM_SETFONT, ByVal m_font.Handle, ByVal MAKELPARAM(False, 0)
End Property

Public Property Let Text(newText As String)

Dim bString() As Byte

    bString = newText
    SendMessageW m_hEdit, WM_SETTEXT, 0&, StrPtr(bString)
    
End Property

Public Property Get Text() As String

Dim iLength As Long
Dim bString As String

    iLength = SendMessageW(m_hEdit, WM_GETTEXTLENGTH, 0&, 0&)
    
    bString = Space$(iLength)
    SendMessageW m_hEdit, WM_GETTEXT, iLength + 1, StrPtr(bString)

    Text = bString

End Property

Public Property Get HasFocus() As Boolean
    HasFocus = m_hasFocus
End Property

Public Property Get hWnd() As Long
    hWnd = m_container.hWnd
End Property

Public Property Get real_hWnd() As Long
    real_hWnd = m_hEdit
End Property

Private Sub Class_Initialize()
    'MsgBox "VISTASEARCHBOX"
    Set m_container = New frmZOrderKeeper
    m_container.SubclassWindow
    
    m_hEdit = CreateWindowExW(0, StrPtr("EDIT"), StrPtr("ViStart_Edit"), _
                WS_VISIBLE Or WS_CHILD Or WS_EX_TOOLWINDOW Or ES_AUTOHSCROLL Or ES_AUTOVSCROLL Or ES_MULTILINE, _
                0, 0, 0, 0, m_container.hWnd, 0, GetModuleHandle(0), 0)
    
    m_hDC = GetWindowDC(m_hEdit)
    
    SetWindowLong m_container.hWnd, GWL_STYLE, WS_VISIBLE Or WS_EX_TOOLWINDOW
    
    m_defWindowProc = HookWindow(m_hEdit, Me)
    
    m_BackColour = vbWhite
    m_ForeColour = vbBlack
    m_FocusColour = RGB(109, 109, 109)
    
    If g_WindowsXP Then
        Set m_backgroundBrush = New GDIBrush
        m_backgroundBrush.Colour = vbWhite

    End If
End Sub

Private Sub Class_Terminate()
    If Not m_container Is Nothing Then Unload m_container
    
    If m_hEdit <> 0 Then
        DestroyWindow m_hEdit
    End If
End Sub

Private Sub WindowsXP_WhiteTextBoxFix()

Dim targetRect As RECT

    targetRect.Left = m_WindowWidth - 3
    targetRect.Right = m_WindowWidth
    targetRect.Bottom = m_WindowHeight
    
    FillRect m_hDC, targetRect, m_backgroundBrush.Value
    'SetPixel m_Hdc, 1, 1, vbBlack
    
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long

    On Error GoTo Handler

    'form-specific handler
    Select Case msg

    Case WM_SETFOCUS
        m_hasFocus = True
        RaiseEvent onFocus
    
    Case WM_KILLFOCUS
        m_hasFocus = False
        RaiseEvent onLostFocus
        
    Case WM_KEYDOWN
        RaiseEvent onKeyDown(wp)
        
    Case WM_KEYUP
        RaiseEvent onKeyUp(wp)
    
        If wp = vbKeyReturn Then
            Exit Function
        End If
        
    Case WM_CHAR
        If wp = vbKeyReturn Then
            Exit Function
        End If
        
    Case WM_SIZE
        'GetWindowRect m_hWnd, m_WindowDimensions
        m_WindowWidth = LOWORD(lp)
        m_WindowHeight = HiWord(lp)
        
    Case WM_PAINT
        
        
        'this takes care of messages when the
        'handle specified is not that of the form
        IHookSink_WindowProc = CallWindowProc(m_defWindowProc, _
                                                    hWnd, _
                                                    msg, _
                                                    wp, _
                                                    lp)
        
        If g_WindowsXP Then
            WindowsXP_WhiteTextBoxFix
        End If
        
    Case WM_MOUSEWHEEL
        RaiseEvent onMouseWheel(wp)
    
    End Select
    
    'Exit Function
Handler:
    'this takes care of messages when the
    'handle specified is not that of the form
    IHookSink_WindowProc = CallWindowProc(m_defWindowProc, _
                                                hWnd, _
                                                msg, _
                                                wp, _
                                                lp)
End Function

Private Sub m_container_onChange(ByVal hWnd As Long)
    If hWnd = m_hEdit Then
        RaiseEvent onChange
    End If
End Sub

Private Sub m_container_onColorEdit(ByVal hWnd As Long, ByVal hEditBox As Long)
    If hWnd = m_hEdit Then
            
        Dim hdc As Long: hdc = hEditBox
        
        SetBkMode hdc, OPAQUE
        SetBkColor hdc, m_BackColour
        
        'SetTextColor hdc, RGB(109, 109, 109)
        If Me.HasFocus Then
            SetTextColor hdc, m_FocusColour
        ElseIf Me.Text = GetPublicString("strStartSearch", "Start Search") Then
            SetTextColor hdc, m_ForeColour
        End If
            
    End If
End Sub

Private Sub m_container_onResize()
    On Error Resume Next
    MoveWindow m_hEdit, 0, 0, m_container.ScaleWidth, m_container.ScaleHeight, True
End Sub
