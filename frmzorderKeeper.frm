VERSION 5.00
Begin VB.Form frmZOrderKeeper 
   BorderStyle     =   0  'None
   Caption         =   "ViStart Container"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   14
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmZOrderKeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event onLoad()
Public Event onResize()

Public Event onColorEdit(ByVal hWnd As Long, ByVal hEditHdc As Long)
Public Event onChange(ByVal hWnd As Long)

Implements IHookSink

Private m_Hooked As Boolean

Private Sub ShutDownApp()
    g_Exiting = True
    
Dim F As Form
    For Each F In Forms
        If F.hWnd <> Me.hWnd Then
            Unload F
        End If
    Next
End Sub

Public Sub InitiateLogOff()
    ShutDownApp
    PowerHelper.ExitWindowsEx EWX_LOGOFF, EWX_FORCEIFHUNG
End Sub

Public Sub InitiateRestart()
    ShutDownApp
    PowerHelper.ExitWindowsEx EWX_REBOOT, EWX_FORCEIFHUNG
End Sub

Public Sub InitiateShutDown()
    ShutDownApp
    PowerHelper.ExitWindowsEx EWX_POWEROFF, EWX_FORCEIFHUNG
End Sub

Private Sub Form_Initialize()
    Me.Move 0, 0, 0, 0
End Sub

Private Sub Form_Load()
    RaiseEvent onLoad
End Sub

Private Sub Form_Resize()
    RaiseEvent onResize
End Sub

Public Sub SubclassWindow()
    Call HookWindow(Me.hWnd, Me)
    m_Hooked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If m_Hooked Then
        UnhookWindow Me.hWnd
    End If
    
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long

    On Error GoTo Handler

    If msg = WM_CTLCOLOREDIT Then
        RaiseEvent onColorEdit(lp, wp)
    ElseIf msg = WM_COMMAND Then
        
        If HiWord(wp) = EN_CHANGE Then
            RaiseEvent onChange(lp)
        End If
        
    Else
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wp, lp)
    End If

    Exit Function
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       InvokeWindowProc(hWnd, msg, wp, lp)
End Function
