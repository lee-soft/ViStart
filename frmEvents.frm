VERSION 5.00
Begin VB.Form frmEvents 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ViStart_Event_Handler"
   ClientHeight    =   2030
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   2270
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2030
   ScaleWidth      =   2270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer timzOrderCheck 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   1440
   End
   Begin VB.Timer timViOrb 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1680
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Force Invoke"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer timHookCheck 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   360
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long

Private Declare Function Shell_NotifyIcon _
  Lib "shell32" Alias "Shell_NotifyIconA" ( _
  ByVal message As Long, Data As NotifyIconData) As Boolean

Dim lng_myTop As Long
Dim lng_myLeft As Long

'Private m_taskbarKnownLeft As Long
'Private m_taskbarKnownTop As Long
Private m_taskbarKnownRect As win.RECT

Dim lngOldStyle As Long
Dim bStartButtonLocked As Boolean

' FormMain.frm - Add an icon to the system tray.
' Copyright (c) 2001. All Rights Reserved.
' By Paul Kimmel. pkimmel@softconcepts.com

' Type passed to Shell_NotifyIcon
Private Type NotifyIconData
  Size As Long
  Handle As Long
  Id As Long
  Flags As Long
  CallBackMessage As Long
  Icon As Long
  Tip As String * 64
End Type

' Constants for managing System Tray tasks, foudn in shellapi.h
Private Const AddIcon = &H0
Private Const DeleteIcon = &H2

Private Const MessageFlag = &H1
Private Const IconFlag = &H2
Private Const TipFlag = &H4

Private WithEvents m_winVistaMenu As frmVistaMenu
Attribute m_winVistaMenu.VB_VarHelpID = -1
Private WithEvents m_taskbarParent As frmZOrderKeeper
Attribute m_taskbarParent.VB_VarHelpID = -1

Private m_startButton As frmStartOrb
Attribute m_startButton.VB_VarHelpID = -1

Private WithEvents m_startButtonEvents As IPngImageEvents
Attribute m_startButtonEvents.VB_VarHelpID = -1
Private WithEvents m_startMenuBase As frmStartMenuBase
Attribute m_startMenuBase.VB_VarHelpID = -1
Private WithEvents m_startOptions As frmVistaMenu
Attribute m_startOptions.VB_VarHelpID = -1
Private WithEvents m_logManager As LogManager
Attribute m_logManager.VB_VarHelpID = -1

Private m_windows8TaskBar As Windows8TaskBar

Private Data As NotifyIconData

Public bInvokeViStart As Boolean
Public bAllowCommand As Boolean

Private bRecievedStartMenu As Boolean

Private lngHwndViOrb As Long
Private bViOrbOpen As Boolean
Private lngLastDraw As Long

Private m_taskbarRect As RECTL
Private m_inTray As Boolean
Private m_showingWinStartMenu As Boolean

Private m_abNormalDPI As Boolean
Private m_userDPI As Long

Private m_ORB_HEIGHT As Long
Private m_logger As SeverityLogger

Property Get Logger()
    Set Logger = m_logger
End Property

Private Sub InitializeMenu()
    Set m_startOptions = New frmVistaMenu

    m_startOptions.AddItem UserVariable("strOptions"), "OPTIONS", False
    m_startOptions.AddItem UserVariable("strOrbNew"), "NEW_IMAGE"
    m_startOptions.AddItem UserVariable("strOrbReset"), "RESET"
    m_startOptions.AddItem ""
    m_startOptions.AddItem UserVariable("strRun"), "RUN"
    m_startOptions.AddItem ""
    m_startOptions.AddItem UserVariable("strSleep"), "STANDBY"
    
    If g_Windows8 Or g_Windows81 Then
        m_startOptions.AddItem ""
        m_startOptions.AddItem "Show Metro", "SHOW_METRO", False
    End If
End Sub

Function InitializeStartButton()

Dim startOrbPath As String
    startOrbPath = sCon_AppDataPath & "_orbs\" & Settings.CurrentOrb

    ShowWindow g_hwndStartButton, SW_HIDE
    'SendMessage g_hwndStartButton, SC_CLOSE
    'SendMessage CLng(g_hwndStartButton), ByVal WM_SYSCOMMAND, ByVal SC_CLOSE, 0
    
    Set m_startButton = frmStartOrb
    m_startButton.SetContextMenu m_startOptions
    
    Set m_startButton.callback = m_startButtonEvents

    If ShellHelper.g_Windows8 Or ShellHelper.g_Windows7 Then
        m_startButton.mode = 2
        
        ShellHelper.UpdateHwnds
        
        If IsWindow(ShellHelper.g_lngHwndViOrbToolbar) = APITRUE Then
            SetParent m_startButton.hWnd, ShellHelper.g_hwndReBarWindow32
        Else
            SetParent m_startButton.hWnd, ShellHelper.g_lnghwndTaskBar
        End If
    Else
        m_startButton.mode = 1
    
        SetOwner m_startButton.hWnd, m_taskbarParent.hWnd
        SetParent m_taskbarParent.hWnd, ShellHelper.g_hwndStartButton
        
        m_taskbarParent.Show
        m_taskbarParent.Move 0, 0, 0, 0
        
        timzOrderCheck.Enabled = True
    End If
    
    If FileExists(startOrbPath) Then
        m_startButton.Path = startOrbPath
    ElseIf FileExists(ResourcesPath & "start_button.png") Then
        m_startButton.Path = ResourcesPath & "start_button.png"
    ElseIf FileExists(sCon_AppDataPath & "_orbs\default.png") Then
                m_startButton.Path = sCon_AppDataPath & "_orbs\default.png"
        ElseIf FileExists(sCon_AppDataPath & "_orbs\Windows 7.png") Then
                m_startButton.Path = sCon_AppDataPath & "_orbs\Windows 7.png"
        ElseIf FileExists(sCon_AppDataPath & "_orbs\start_button.png") Then
                m_startButton.Path = sCon_AppDataPath & "_orbs\start_button.png"
        ElseIf FileExists(sCon_AppDataPath & "start_button.png") Then
                m_startButton.Path = sCon_AppDataPath & "start_button.png"
    End If
 
    
    m_ORB_HEIGHT = m_startButton.ScaleHeight
        
    Call SetWindowPos(m_startButton.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    
    MoveOrbIfNotOverStartButton
    
    m_startButton.Show
    
    timViOrb.Enabled = True
    bViOrbOpen = True
    
    InitializeStartButton = True
End Function

Private Sub Form_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)
    Set m_logManager = LogManager
End Sub

Private Sub Form_Load()
    InitializeMenu

    'Dont intialize startbutton here because it appears in task-manager
    Set m_startButtonEvents = New IPngImageEvents
    
    'MsgBox "DEBUG::FRMEVENTS"
    Set m_taskbarParent = New frmZOrderKeeper
    Set m_startMenuBase = frmStartMenuBase
    Load m_startMenuBase
    
    m_startMenuBase.SetContextMenu m_startOptions
    
    m_taskbarKnownRect.Top = -1
    m_taskbarKnownRect.Left = -1

    Load m_taskbarParent
    m_taskbarParent.SubclassWindow
    
    If InitializeStartButton = False Then
        ExitApplication
        Exit Sub
    End If
    
    timHookCheck.Enabled = True
    
    If sVar_bDebugMode Then
        Visible = True
    Else
        Visible = False
        
    End If
    
    Set m_winVistaMenu = frmVistaMenu

    bAllowCommand = False
    bInvokeViStart = True
    
    If Settings.EnableTrayIcon Then
        AddIconToTray
    End If
    
    If (g_Windows8 = True) And (g_Windows81 = False) And IsWindow(ShellHelper.g_lngHwndViOrbToolbar) = APIFALSE Then
        Set m_windows8TaskBar = New Windows8TaskBar
    End If

    m_userDPI = GeneralHelper.CurrentDPI

    If m_userDPI >= 144 Then
        m_abNormalDPI = True
    End If
End Sub

Sub AddIconToTray()
    If m_inTray Then
        Exit Sub
    End If

    m_inTray = True

    Data.Size = Len(Data)
    Data.Handle = hWnd
    Data.Id = App.hInstance
    Data.Flags = IconFlag Or TipFlag Or MessageFlag
    Data.CallBackMessage = WM_MOUSEMOVE
    Data.Icon = Icon
    Data.Tip = App.Title & vbNullChar
    Call Shell_NotifyIcon(AddIcon, Data)

End Sub

Sub DeleteIconFromTray()
    m_inTray = False
    Call Shell_NotifyIcon(DeleteIcon, Data)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim message As Long

    message = X / Screen.TwipsPerPixelX

  Select Case message
  
    Case WM_RBUTTONUP
        m_winVistaMenu.Resurrect True, Me
        
  End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_windows8TaskBar = Nothing
    
    If Not g_ViGlanceOpen Then
        ShowWindow ShellHelper.g_hwndStartButton, SW_SHOW
    End If
    
    Unload m_startButton
    Unload m_taskbarParent
    Unload m_winVistaMenu
    Unload m_startMenuBase
    
    DeleteIconFromTray
    ExitApplication
End Sub

Private Sub m_logManager_LogEvent(ByVal level As LogLevel, ByVal Source As String, ByVal message As String, arguments() As String)
    If level = FatalLevel Then
        MsgBox Source & vbCrLf & vbCrLf & message, vbCritical, "Fatal Error"
    End If
End Sub

Private Sub m_startButtonEvents_onMouseDown(ObjSender As Object)

    If lngLastDraw <> 2 Then
        
        If Settings.StartButtonShowsWindowsMenu Then
            If Not m_showingWinStartMenu Then
                'SendMessage g_lnghwndTaskBar, ByVal WM_SYSCOMMAND, ByVal SC_TASKLIST, ByVal 0
                ShowNormalWindowsMenu
            Else
                m_startButton.Show
            End If
        Else
            ActivateStartMenu
        End If
    Else
        lngLastDraw = 0
        m_startButton.DrawPart 0
        m_startMenuBase.CloseMe
        
    End If

End Sub

Private Sub m_startMenuBase_onClose()
    lngLastDraw = -1
    CheckRolloverStatus_ViStart
End Sub

Private Sub m_startMenuBase_onRequestNewResize()
    m_taskbarKnownRect.Bottom = -1
    m_taskbarKnownRect.Left = -1
    m_taskbarKnownRect.Right = -1
    m_taskbarKnownRect.Top = -1
End Sub

Private Sub m_startMenuBase_onSkinChange()
    'if user has no customized orb, then refresh orb from current skin
    If Settings.CurrentOrb = vbNullString Then
        m_startButton.ResetOrb
    End If
End Sub

Private Sub m_startOptions_onClick(theItemTag As String)
    On Error GoTo Handler

Dim fileNameHolder As String

    m_startOptions.Hide
    
    m_startMenuBase.CloseMe

    If theItemTag = "NEW_IMAGE" Then
    
        m_startButton.PromptUserSelectNewOrb
        
    ElseIf theItemTag = "RUN" Then
        
        'MsgBox "DEBUG::RUN"
        Dim dummyWin As New frmZOrderKeeper
        dummyWin.Move m_startButton.Left, m_startButton.Top, 0, 0
        
        ShowRun dummyWin.hWnd
        Unload dummyWin
        
    ElseIf theItemTag = "RESET" Then
        
        m_startButton.ResetOrb
        
    ElseIf theItemTag = "STANDBY" Then
    
        SetSuspendState False, True, False
    
    ElseIf theItemTag = "SHOW_METRO" Then
    
        ShowNormalWindowsMenu
        
    ElseIf theItemTag = "OPTIONS" Then
    
        frmControlPanel.Show
        'Set m_changeSkin = New frmSkinSelect
        'm_changeSkin.Show
    
    End If
    
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, "I/O Error"
End Sub

Private Sub m_startOptions_onInActive()
    m_startOptions.Hide
End Sub

Private Sub m_taskbarParent_onLoad()
    ShellHelper.UpdateHwnds
    

End Sub

Private Sub m_winVistaMenu_onInActive()
    'm_winVistaMenu_Subclass.UnSubClass
    m_winVistaMenu.Hide
End Sub

Sub ActivateSearchText(ByVal KeyCode As Long)
    'If keyCode = vbKeyReturn Then
    
    m_startMenuBase.ActivateSearchText
End Sub

Sub ActivateStartMenu()
Dim TaskBar As win.RECT
Dim taskbarEdge As AbeBarEnum

    TaskBar = GetTaskBarPosition
    
    lngLastDraw = 2
    m_startButton.DrawPart 2
    
    If IsRectDifferent(m_taskbarKnownRect, TaskBar) Then
        'Taskbar must have moved
        Logger.Trace "Taskbar dimensions have changed", "ActivateStartMenu"
        
        taskbarEdge = GetTaskBarEdge()
        m_taskbarKnownRect = TaskBar
        
        If taskbarEdge = abe_bottom Then
            lng_myLeft = Layout.XOffset
            
            If m_abNormalDPI Then
                lng_myTop = CalculateTopBasedOnDPI(m_userDPI, m_startMenuBase.ScaleHeight)
            Else
                lng_myTop = (TaskBar.Top - m_startMenuBase.ScaleHeight) - Layout.YOffset
            End If
        ElseIf taskbarEdge = ABE_TOP Then
            lng_myLeft = Layout.XOffset
            lng_myTop = (TaskBar.Bottom) + Layout.YOffset
        ElseIf taskbarEdge = ABE_LEFT Then
            lng_myLeft = Layout.XOffset
            lng_myTop = 55 - Layout.YOffset
        ElseIf taskbarEdge = ABE_RIGHT Then
            lng_myLeft = (TaskBar.Right - m_startMenuBase.ScaleWidth) - Layout.XOffset
            lng_myTop = 55 - Layout.YOffset
        End If
    End If
    
    'Transparent treeview repaint bug
    SetForegroundWindow m_startMenuBase.hWnd

    'Push StartMenu in front of TaskBar
    SetWindowPos m_startMenuBase.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

    With m_startMenuBase
        If m_startMenuBase.BlurEnabled Then
        
            .UpdateDesktopImage CreatePointL(lng_myTop, lng_myLeft)
            .ReDraw
        End If
    
        '.Left = lng_myLeft * Screen.TwipsPerPixelX
        '.Top = lng_myTop * Screen.TwipsPerPixelY
        MoveWindow .hWnd, lng_myLeft, lng_myTop, .ScaleWidth, .ScaleHeight, APITRUE
    
        '.AlignElements
        '.AlignHotSpots
        
        
        .ShowMe
        win.SetFocus .hWnd
        
        
     End With
End Sub

Sub Test()

    With m_startMenuBase
        .Visible = False
        .Left = lng_myLeft
        .Top = lng_myTop
            
        '.AlignElements
        
        .Top = -9600
        .Visible = True
     End With

End Sub

Public Function CheckRolloverStatus_WindowsStartMenu()

    If MouseInsideWindow(m_startButton.hWnd) Then

        If win.IsWindowVisible(g_lnghwndStartMenu) = APIFALSE And frmStartMenuBase.Visible = False Then
            If lngLastDraw <> 1 Then
                lngLastDraw = 1
                m_startButton.DrawPart 1
            End If
            
        ElseIf frmStartMenuBase.Visible = True Or win.IsWindowVisible(g_lnghwndStartMenu) = APITRUE Then
             If lngLastDraw <> 2 Then
                lngLastDraw = 2
                m_startButton.DrawPart 2
            End If
        End If
        
    Else
    
        If win.IsWindowVisible(g_lnghwndStartMenu) = APIFALSE And frmStartMenuBase.Visible = False Then
            If lngLastDraw <> 0 Then
                lngLastDraw = 0
                m_startButton.DrawPart 0
            End If
        End If
    End If

End Function

Public Function CheckRolloverStatus_ViStart()

    If lngLastDraw <> 2 Then
        If MouseInsideWindow(m_startButton.hWnd) Then
            
            If lngLastDraw <> 1 Then
                lngLastDraw = 1
                m_startButton.DrawPart 1
            End If
        Else
            If lngLastDraw <> 0 Then
                lngLastDraw = 0
                m_startButton.DrawPart 0
            End If
        End If
    End If
    
End Function

Private Function GetHeightWithRegardsToFix(theHeight As Long) As Long

    If MainHelper.g_viOrb_fullHeight Then
        GetHeightWithRegardsToFix = theHeight
    Else
        GetHeightWithRegardsToFix = theHeight - 1
    End If

End Function

Public Function MoveOrbIfNotOverStartButton()
    If m_startButton.WindowState <> vbNormal Then
        Exit Function
    End If

Dim recViOrbToolbar As RECT
Dim recReBar32 As RECT
Dim recOrb As RECT

Dim lngTop As Long
Dim lngLeft As Long

Dim taskbarEdge As AbeBarEnum
Dim taskBarHeight As Long
Dim taskBarWidth As Long

Dim recStartButton As RECT
Dim buttonMaxHeight As Long
Dim topValueSet As Boolean: topValueSet = False

    taskbarEdge = GetTaskBarEdge()

    GetWindowRect IIf(g_Windows11, g_lnghwndTaskBar, g_hwndReBarWindow32), recReBar32
    'GetWindowRect g_hwndReBarWindow32, recReBar32
    GetWindowRect m_startButton.hWnd, recOrb
    
    taskBarHeight = (recReBar32.Bottom - recReBar32.Top)
    taskBarWidth = (recReBar32.Right - recReBar32.Left)

    If taskbarEdge = ABE_LEFT Then

        If Not g_Windows81 Then
            buttonMaxHeight = 60
        Else
            topValueSet = True
            lngTop = 0
            
            buttonMaxHeight = 39
            lngLeft = (taskBarWidth / 2) - (m_startButton.ScaleWidth / 2)
        End If
         
        If m_startButton.MaxHeight <> buttonMaxHeight Then
            m_startButton.MaxHeight = buttonMaxHeight
            m_ORB_HEIGHT = m_startButton.ScaleHeight
        End If
         
    ElseIf taskbarEdge = ABE_RIGHT Then

        If Not g_Windows81 Then
            buttonMaxHeight = 60
        Else
            topValueSet = True
            lngTop = 0
            
            buttonMaxHeight = 39
            lngLeft = (taskBarWidth / 2) - (m_startButton.ScaleWidth / 2)
        End If
        
        If m_startButton.MaxHeight <> buttonMaxHeight Then
            m_startButton.MaxHeight = buttonMaxHeight
            m_ORB_HEIGHT = m_startButton.ScaleHeight
        End If
    
    ElseIf taskbarEdge = ABE_TOP Then
    
        If m_startButton.MaxHeight <> GetHeightWithRegardsToFix(taskBarHeight) Then
            m_startButton.MaxHeight = GetHeightWithRegardsToFix(taskBarHeight)
            m_ORB_HEIGHT = m_startButton.ScaleHeight
        End If
    
        topValueSet = True
        lngTop = (((m_ORB_HEIGHT) / 2) - (taskBarHeight) / 2) * -1
        'Windows taskbar has a border of 2 pixels
        If Not MainHelper.g_viOrb_fullHeight Then
            lngTop = lngTop + 2
        End If

        If g_Windows81 Then
            lngLeft = m_startButton.Windows81LeftDifference
            
            topValueSet = True
            lngTop = lngTop - 1
            
            'Prevent weird bug of windows 8 placing orb in middle of taskbar
            If lngTop < 0 Then
                lngTop = 0
            End If
        End If
            
    ElseIf taskbarEdge = abe_bottom Then
    
        'MsgBox "abe_bottom"
    
        If m_startButton.MaxHeight <> GetHeightWithRegardsToFix(taskBarHeight) Then
            m_startButton.MaxHeight = GetHeightWithRegardsToFix(taskBarHeight)
            m_ORB_HEIGHT = m_startButton.ScaleHeight
        End If
        
        'MsgBox "taskbarheight " & taskBarHeight
        
        topValueSet = True
        lngTop = (((m_ORB_HEIGHT) / 2) - (taskBarHeight) / 2) * -1

        'Windows taskbar has a border of 2 pixels
        If Not MainHelper.g_viOrb_fullHeight Then
            lngTop = lngTop + 2
        End If
        
        If g_Windows81 Then
            lngLeft = m_startButton.Windows81LeftDifference
            lngTop = lngTop - 1
        
            'Prevent weird bug of windows 8 placing orb in middle of taskbar
            If lngTop < 0 Then
                lngTop = 0
            End If
        End If

    End If

    If topValueSet Then
        'TO BE FIXED!
        
        If m_startButton.mode = 1 And Not g_Windows11 Then
            'Windows XP and Vista and 11?
            
            If ((recStartButton.Left) <> (recOrb.Left) Or _
                (recReBar32.Top + lngTop) <> recOrb.Top) Then

                MoveWindow m_startButton.hWnd, recStartButton.Left, recReBar32.Top + lngTop, m_startButton.ScaleWidth, m_startButton.ScaleHeight, True
            End If
                        
        ElseIf g_Windows11 Then
            ' Windows 11+
            MoveWindow m_startButton.hWnd, recStartButton.Left + 2, recReBar32.Top + lngTop, m_startButton.ScaleWidth, m_startButton.ScaleHeight, True

        Else
        
                
            'Windows 7 and 8
            If IsWindow(g_lngHwndViOrbToolbar) = APITRUE Then
                
                Dim newPoint As POINTL
                
                GetClientRect g_lngHwndViOrbToolbar, recViOrbToolbar
                newPoint.X = recViOrbToolbar.Left
                newPoint.Y = recViOrbToolbar.Top
                MapWindowPoints g_lngHwndViOrbToolbar, CLng(getParent(g_lngHwndViOrbToolbar)), newPoint, 2
                
                MoveWindow m_startButton.hWnd, newPoint.X - 40, lngTop, m_startButton.ScaleWidth, m_startButton.ScaleHeight, True
                
            Else
                If (recReBar32.Top + lngTop) <> (recOrb.Top) Then
            
                    MoveWindow m_startButton.hWnd, lngLeft, lngTop, m_startButton.ScaleWidth, m_startButton.ScaleHeight, True
                End If
            End If
        End If
    Else
        If m_startButton.mode = 1 Then
            If ((recStartButton.Left) <> (recOrb.Left) Or _
                (recStartButton.Top - IIf(g_WindowsVista, 0, 5)) <> recOrb.Top) Then
                
                MoveWindow m_startButton.hWnd, recStartButton.Left, recStartButton.Top - IIf(g_WindowsVista, 0, 5), m_startButton.ScaleWidth, m_startButton.ScaleHeight, True
            End If
        Else
            lngLeft = (taskBarWidth / 2) - (m_startButton.ScaleWidth / 2)
            lngTop = 1
        
            If (recReBar32.Left + lngLeft) <> (recOrb.Left) Then
                MoveWindow m_startButton.hWnd, lngLeft, lngTop, m_startButton.ScaleWidth, m_startButton.ScaleHeight, True
            End If
        End If
    End If
End Function

Private Sub timHookCheck_Timer()

    'MoveOrbIfNotOverStartButton

    If ShellHelper.UpdateHwnds Then
        timHookCheck.Interval = 5000
    End If
    
    If bViOrbOpen Then
        If g_ViGlanceOpen Then
            bViOrbOpen = False
            
            m_startButton.Caption = "##VIGLANCE_MODE##"
            m_startButton.Hide
        End If
    Else
        If Not g_ViGlanceOpen Then
            bViOrbOpen = True
            
            m_startButton.Caption = "##VISTART_MODE##"
            ShowWindow m_startButton.hWnd, SW_SHOWNOACTIVATE
        End If
    End If


End Sub

Private Sub timViOrb_Timer()

    If IsWindowVisible(g_hwndStartButton) = APITRUE Then
        Logger.Trace "Closing Start Button!", "timViOrb_Timer"
        
        ShowWindow g_hwndStartButton, SW_HIDE
        Call SetWindowPos(m_startButton.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If

    If bViOrbOpen Then
    
        If Settings.StartButtonShowsWindowsMenu Then
            CheckRolloverStatus_WindowsStartMenu
        Else
            CheckRolloverStatus_ViStart
        End If

        
        MoveOrbIfNotOverStartButton
        
        If Not m_windows8TaskBar Is Nothing Then
            m_windows8TaskBar.CheckTaskbarSize
        End If
    
        If Settings.StartButtonShowsWindowsMenu Then
            If IsWindowVisible(g_lnghwndStartMenu) = APITRUE Then
                m_showingWinStartMenu = True
            Else
                m_showingWinStartMenu = False
            End If
            
            
        End If
    End If
End Sub

Private Sub timzOrderCheck_Timer()

Dim hWndForeGroundWindow As Long
Dim zOrderOrb As Long
Static m_NotTopMost As Boolean
    
    hWndForeGroundWindow = GetForegroundWindow
    zOrderOrb = GetZOrder(m_startButton.hWnd)

    If (GetZOrder(g_lnghwndTaskBar) < zOrderOrb) And m_NotTopMost = False Then
        SetWindowPos m_startButton.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        If Not hWndBelongToUs(hWndForeGroundWindow) And _
                    hWndForeGroundWindow <> g_lnghwndTaskBar Then
            
            'taskbar behind target window
            If IsTaskBarBehindWindow(hWndForeGroundWindow) Then
                If IsWindowTopMost(hWndForeGroundWindow) = False Then
                    m_NotTopMost = True
                    m_startButton.Hide
                Else
                    'topmost target Window is behind Orb
                    If zOrderOrb < GetZOrder(hWndForeGroundWindow) Then
                        SetWindowPos hWndForeGroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                        m_NotTopMost = True
                    End If
                End If
            Else
                If m_NotTopMost = True Then
                    m_NotTopMost = False
                    
                    Logger.Trace "Showing ViStart start button", "timzOrderCheck_Timer"
                    ShowWindow m_startButton.hWnd, SW_SHOWNOACTIVATE
                End If
            End If
        End If
    End If
End Sub
