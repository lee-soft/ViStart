VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Windows8Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CharmsBarTop_Disabled As Boolean
Public CharmsBarBottom_Disabled As Boolean
Public DragToClose_Disabled As Boolean
Public WindowsStartCorner_Disabled As Boolean
Public AppSwitcher_Disabled As Boolean
Public SkipMetro_Enabled As Boolean

Private m_desktopShellWindow As Long
Private m_iniFile As WinProfile

Private m_logger As SeverityLogger

Property Get Logger()
    Set Logger = m_logger
End Property

Public Function IsMetroRunning() As Boolean

    If FindWindow("ApplicationManager_DesktopShellWindow", vbNullString) <> 0 Then
        If FindWindow("EdgeUiInputWndClass", "") <> 0 Then
            IsMetroRunning = True
        End If
    End If

End Function

Private Function GetDesktopShellWindow()

Dim loopCount As Long

    Do
        m_desktopShellWindow = FindWindow("ApplicationManager_DesktopShellWindow", vbNullString)
        
        If loopCount = 10 Then
            Logger.Trace "Failed to aquire desktop shell window"
            Exit Do
        End If
        
        If m_desktopShellWindow = 0 Then Sleep 1000
        loopCount = loopCount + 1
        
        Logger.Trace "Continue loop [GetDesktopShellWindow]"
        
    Loop While (m_desktopShellWindow = 0)

End Function

Public Function ActionSettings()
    Logger.Trace "Actioning Settings!"

    GetDesktopShellWindow
    If m_desktopShellWindow = 0 Then
        Exit Function
    End If

    If SkipMetro_Enabled Then
        SkipMetroScreen
    End If

    If CharmsBarTop_Disabled Or CharmsBarBottom_Disabled Or DragToClose_Disabled Or WindowsStartCorner_Disabled Or AppSwitcher_Disabled Or SkipMetro_Enabled Then
        DisableMetroShit
    End If
End Function

Private Function IsDesktopShowing(ByRef hWnd As Long)
    
End Function

Public Function SkipMetroScreen()

Dim desktopShellWindow_hWnd As Long
Dim loopCount As Long
Dim nextForegroundWindow As Long

    desktopShellWindow_hWnd = m_desktopShellWindow
    Logger.Trace "desktopShellWindow_hWnd is " & desktopShellWindow_hWnd

    Do
        DoEvents
        
        If desktopShellWindow_hWnd <> 0 Then
            nextForegroundWindow = GetForegroundWindow()
            
            BringWindowToTop desktopShellWindow_hWnd
            SetForegroundWindow desktopShellWindow_hWnd
            
            nextForegroundWindow = GetForegroundWindow()
            
            If nextForegroundWindow = desktopShellWindow_hWnd Then
                Logger.Trace "Succesfully brought to front!"
                Exit Do
            Else
                Logger.Trace "Failed to bring to front, minimizing foreground window"
                nextForegroundWindow = GetForegroundWindow()

                If GetWindowClassNameByHwnd(nextForegroundWindow) = "ImmersiveLauncher" Then
                    
                    ShowWindow nextForegroundWindow, SW_MINIMIZE
                    
                    BringWindowToTop desktopShellWindow_hWnd
                    SetForegroundWindow desktopShellWindow_hWnd
                    
                    ShowWindow nextForegroundWindow, SW_RESTORE
                    
                    If GetForegroundWindow() = desktopShellWindow_hWnd Then
                        Logger.Trace "Succesfully brought to front!"
                        Exit Do
                    End If
                    
                    nextForegroundWindow = GetForegroundWindow()
                End If
            End If
        End If
        
        desktopShellWindow_hWnd = FindWindow("ApplicationManager_DesktopShellWindow", vbNullString)
        loopCount = loopCount + 1
        
        If loopCount = 10 Then
            Logger.Trace "Failed to bring to front!"
            Exit Do
        Else
            'Send WinKey+M to minimize all windows
            keybd_event VK_LWIN, 0, 0, 0
            keybd_event VK_D, 0, 0, 0
            keybd_event VK_D, 0, KEYEVENTF_KEYUP, 0
            keybd_event VK_LWIN, 0, KEYEVENTF_KEYUP, 0
        
            Sleep 10
            Logger.Trace "Continue loop [SkipMetroScreen]"
        End If
    Loop
    
End Function

Sub DisableMetroShit()

Dim desktopShellWindow_hWnd As Long
    desktopShellWindow_hWnd = m_desktopShellWindow

    If desktopShellWindow_hWnd = 0 Then Exit Sub

Dim programManager As Long: programManager = -1
Dim foregroundWindow As Long: foregroundWindow = -2

Dim num2 As Long
Dim num1 As Long
Dim num As Long

Dim myThread As Long

Dim appSwitchhWnds As Collection
Dim startButtonhWnds As Collection
Dim charmsBarTophWnds As Collection
Dim charmsBarBottomhWnds As Collection
Dim dragToClosehWnds As Collection

Dim edgeUiInputWndClasshWnd As Long
Dim windowRect As RECT
Dim identifiedEdgeType As Boolean
Dim loopCount As Long


    Logger.Trace "Starting to disable modern UI stuff "

    Set appSwitchhWnds = New Collection
    Set startButtonhWnds = New Collection
    Set charmsBarTophWnds = New Collection
    Set charmsBarBottomhWnds = New Collection
    Set dragToClosehWnds = New Collection
    
    loopCount = 0
    
    Do
        DoEvents
        Sleep 1000
        edgeUiInputWndClasshWnd = FindWindow("EdgeUiInputWndClass", "")
        
        If loopCount = 5 Then
            Logger.Trace "Failed to aquire EdgeUiInputWndClass hWnd, aborting.. perhaps operation is done already?"
            Exit Sub
        End If
        
        loopCount = loopCount + 1
        
    Loop While (edgeUiInputWndClasshWnd = 0)
    Sleep 1000
    
    While (edgeUiInputWndClasshWnd <> 0)
        DoEvents
        GetWindowRect edgeUiInputWndClasshWnd, windowRect
        
        num = Math.Abs(windowRect.Left - windowRect.Right)
        num1 = Math.Abs(windowRect.Top - windowRect.Bottom)
        identifiedEdgeType = False
        
        If (windowRect.Left = 0 And windowRect.Top = 0 And num1 > 0 And num = num1) Then
            appSwitchhWnds.Add edgeUiInputWndClasshWnd
            identifiedEdgeType = True
        End If
        
        If (windowRect.Left = 0 And windowRect.Top > 0 And num1 > 0 And num = num1) Then
            startButtonhWnds.Add edgeUiInputWndClasshWnd
            identifiedEdgeType = True
        End If
        
        If (windowRect.Left > 0 And windowRect.Top = 0 And num1 > 0 And num = num1) Then
            charmsBarTophWnds.Add edgeUiInputWndClasshWnd
            identifiedEdgeType = True
        End If
        
        If (windowRect.Left > 0 And windowRect.Top > 0 And num1 > 0 And num = num1) Then
            charmsBarBottomhWnds.Add edgeUiInputWndClasshWnd
            identifiedEdgeType = True
        End If

        If (Not identifiedEdgeType) Then
            dragToClosehWnds.Add edgeUiInputWndClasshWnd
        End If
        
        SetWindowText edgeUiInputWndClasshWnd, "EdgeUiInputWndClass_Charted"
        edgeUiInputWndClasshWnd = FindWindow("EdgeUiInputWndClass", "")
        
    Wend
    
    If WindowsStartCorner_Disabled Then
        Logger.Trace "Closing Start Button hWnds"
        ClosehWndList startButtonhWnds
    End If
    
    If AppSwitcher_Disabled Then
        Logger.Trace "Closing App Switcher hWnds"
        ClosehWndList appSwitchhWnds
    End If
    
    If CharmsBarTop_Disabled Then
        Logger.Trace "Closing CharmsBar Top hWnds"
        ClosehWndList charmsBarTophWnds
    End If
    
    If CharmsBarBottom_Disabled Then
        Logger.Trace "Closing CharmsBar Bottom hWnds"
        ClosehWndList charmsBarBottomhWnds
    End If
    
    If DragToClose_Disabled Then
        Logger.Trace "Closing DragToClose hWnds"
        ClosehWndList dragToClosehWnds
    End If
    
    DoEvents
    
    Logger.Trace "Finished!"
    Unload frmNukeMetro

End Sub

Private Sub ClosehWndList(ByRef sourceList As Collection)

Dim listIndex As Long
Dim thishWnd As Long

    For listIndex = 1 To sourceList.count
        thishWnd = sourceList(listIndex)
        SendMessage CLng(thishWnd), ByVal WM_SYSCOMMAND, ByVal SC_CLOSE, 0
    Next

End Sub

Public Function ReadSettings()
    'CharmsBarTop_Disabled = Registry.RegObj.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_charmsbar_top", False)
    'CharmsBarBottom_Disabled = Registry.RegObj.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_charmsbar_bottom", False)
    'DragToClose_Disabled = Registry.RegObj.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_drag_to_close", False)
    'WindowsStartCorner_Disabled = Registry.RegObj.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_windows_start_corner", False)
    'AppSwitcher_Disabled = Registry.RegObj.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_app_switcher", False)
    'SkipMetro_Enabled = Registry.RegObj.GetBooleanValue(HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "enable_skip_metro", False)

    CharmsBarTop_Disabled = CBool(m_iniFile.ReadINIValue("metro", "disable_charmsbar_top", False))
    CharmsBarBottom_Disabled = CBool(m_iniFile.ReadINIValue("metro", "disable_charmsbar_bottom", False))
    DragToClose_Disabled = CBool(m_iniFile.ReadINIValue("metro", "disable_drag_to_close", False))
    WindowsStartCorner_Disabled = CBool(m_iniFile.ReadINIValue("metro", "disable_windows_start_corner", False))
    AppSwitcher_Disabled = CBool(m_iniFile.ReadINIValue("metro", "disable_app_switcher", False))
    SkipMetro_Enabled = CBool(m_iniFile.ReadINIValue("metro", "enable_skip_metro", False))
End Function

Public Function DumpOptions() As Boolean
    'Registry.RegObj.CreateKey HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", True

    'Registry.RegObj.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_charmsbar_top", CharmsBarTop_Disabled
    'Registry.RegObj.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_charmsbar_bottom", CharmsBarBottom_Disabled
    'Registry.RegObj.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_drag_to_close", DragToClose_Disabled
    'Registry.RegObj.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_windows_start_corner", WindowsStartCorner_Disabled
    'Registry.RegObj.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "disable_app_switcher", AppSwitcher_Disabled
    'Registry.RegObj.SetBooleanValue HKEY_LOCAL_MACHINE, "Software\ViStart\Windows8Utility", "enable_skip_metro", SkipMetro_Enabled

    m_iniFile.WriteINIValue "metro", "disable_charmsbar_top", CharmsBarTop_Disabled
    m_iniFile.WriteINIValue "metro", "disable_charmsbar_bottom", CharmsBarBottom_Disabled
    m_iniFile.WriteINIValue "metro", "disable_drag_to_close", DragToClose_Disabled
    m_iniFile.WriteINIValue "metro", "disable_windows_start_corner", WindowsStartCorner_Disabled
    m_iniFile.WriteINIValue "metro", "disable_app_switcher", AppSwitcher_Disabled
    m_iniFile.WriteINIValue "metro", "enable_skip_metro", SkipMetro_Enabled
End Function

Private Sub Class_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)
    Set m_iniFile = New WinProfile
    
    m_iniFile.INIPath = Environ$("appdata") & "\ViStart\nuke_metro.ini"
    ReadSettings
End Sub

Private Sub Class_Terminate()
    DumpOptions
End Sub
