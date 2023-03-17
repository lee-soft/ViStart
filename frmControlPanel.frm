VERSION 5.00
Begin VB.Form frmControlPanel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Panel"
   ClientHeight    =   11955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19095
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControlPanel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   797
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1273
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Index           =   3
      Left            =   5140
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   577
      TabIndex        =   21
      Top             =   920
      Visible         =   0   'False
      Width           =   8650
      Begin ViStart.VMLDocument VMLViewer 
         Height          =   8175
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   8655
         _extentx        =   15266
         _extenty        =   12991
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Index           =   1
      Left            =   6300
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   572
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   8580
      Begin VB.PictureBox picClient_1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   600
         ScaleHeight     =   313
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   517
         TabIndex        =   13
         Top             =   240
         Width           =   7755
         Begin VB.CommandButton cmdRestoreSpecialFolders 
            Caption         =   "Restore Defaults"
            Height          =   495
            Left            =   2460
            TabIndex        =   44
            Top             =   3020
            Width           =   2565
         End
         Begin VB.CheckBox chkUserPicture 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show user picture"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   3615
         End
         Begin VB.CheckBox chkProgramMenuFirst 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show program menu first"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   3855
         End
         Begin ViStart.SettingsOption MenuItem 
            Height          =   615
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   4110
            Visible         =   0   'False
            Width           =   6735
            _extentx        =   0
            _extenty        =   0
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmControlPanel.frx":74F2
            Height          =   555
            Left            =   150
            TabIndex        =   43
            Top             =   2250
            Width           =   7500
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Default settings for Start menu items"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   150
            TabIndex        =   41
            Top             =   3930
            Width           =   7215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Visibility settings"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   7215
         End
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Index           =   2
      Left            =   3480
      ScaleHeight     =   8175
      ScaleWidth      =   8655
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CheckBox chkShowSplashScreen 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show splash screen on startup"
         Height          =   375
         Left            =   720
         TabIndex        =   38
         Top             =   3410
         Width           =   5055
      End
      Begin VB.CommandButton cmdShowMetroShortcut 
         Caption         =   "Restore Windows Start Menu Shortcut"
         Height          =   400
         Left            =   720
         TabIndex        =   37
         Top             =   3890
         Width           =   4935
      End
      Begin VB.PictureBox picWindows8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   480
         ScaleHeight     =   4455
         ScaleWidth      =   8055
         TabIndex        =   29
         Top             =   4320
         Visible         =   0   'False
         Width           =   8055
         Begin VB.CheckBox chkSkipMetroScreen 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Automatically go to desktop when I log in "
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   2880
            Width           =   5055
         End
         Begin VB.CheckBox chkDisableBottomLeftCorner 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable bottom left (Start) hot corner"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   2400
            Width           =   5055
         End
         Begin VB.CheckBox chkDisableDragToClose 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Drag to close"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1920
            Width           =   5055
         End
         Begin VB.CheckBox chkDisableCharmsBar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable CharmsBar"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   1440
            Width           =   5055
         End
         Begin VB.CheckBox chkHotCorners 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable all Windows 8 hot corners"
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   960
            Width           =   5055
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Windows 8 related features require a restart to take effect "
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   840
            TabIndex        =   36
            Top             =   3480
            Width           =   6735
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "How should the Windows 8 features work?"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   7695
         End
      End
      Begin VB.CheckBox chkStartWithWindows 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start ViStart with Windows"
         Height          =   375
         Left            =   720
         TabIndex        =   28
         Top             =   2930
         Width           =   5055
      End
      Begin VB.Timer timDelayUnload 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   6960
         Top             =   4320
      End
      Begin VB.CheckBox chkSystemTray 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show ViStart on the system tray menu"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   2450
         Width           =   5055
      End
      Begin VB.ComboBox cmbWindowsOrb 
         Height          =   360
         ItemData        =   "frmControlPanel.frx":75A3
         Left            =   720
         List            =   "frmControlPanel.frx":75AA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   4935
      End
      Begin VB.ComboBox cmbWindowsKey 
         Height          =   360
         ItemData        =   "frmControlPanel.frx":75BD
         Left            =   720
         List            =   "frmControlPanel.frx":75C4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Set default desktop actions"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   7695
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Index           =   0
      Left            =   3360
      ScaleHeight     =   8175
      ScaleWidth      =   8655
      TabIndex        =   5
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox cmbChildThemes 
         Height          =   360
         ItemData        =   "frmControlPanel.frx":75D7
         Left            =   720
         List            =   "frmControlPanel.frx":75D9
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1800
         Width           =   4935
      End
      Begin VB.ComboBox cmbRollover 
         Height          =   360
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   5520
         Width           =   4935
      End
      Begin VB.CommandButton cmdMoreOrbs 
         Caption         =   "More ..."
         Height          =   495
         Left            =   5880
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdMoreThemes 
         Caption         =   "More..."
         Height          =   495
         Left            =   5880
         TabIndex        =   19
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdPickImage 
         Caption         =   "Pick image...&"
         Height          =   400
         Left            =   5880
         TabIndex        =   9
         Top             =   3345
         Width           =   1575
      End
      Begin VB.CommandButton cmdInstallTheme 
         Caption         =   "Install...&"
         Height          =   400
         Left            =   5880
         TabIndex        =   8
         Top             =   1185
         Width           =   1575
      End
      Begin VB.ComboBox cmbStartOrbs 
         Height          =   360
         ItemData        =   "frmControlPanel.frx":75DB
         Left            =   720
         List            =   "frmControlPanel.frx":75DD
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3360
         Width           =   4935
      End
      Begin VB.ComboBox cmbThemes 
         Height          =   360
         ItemData        =   "frmControlPanel.frx":75DF
         Left            =   720
         List            =   "frmControlPanel.frx":75E1
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   4935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Rollover Skin"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   39
         Top             =   4800
         Width           =   5295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Orb Skin"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   11
         Top             =   2640
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Menu Skin"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   10
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.Timer timReloadAbout 
      Interval        =   1
      Left            =   2760
      Top             =   9720
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Index           =   5
      Left            =   9980
      ScaleHeight     =   8175
      ScaleWidth      =   8655
      TabIndex        =   22
      Top             =   7340
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdStartService 
         Caption         =   "&Start Service"
         Height          =   495
         Left            =   4800
         TabIndex        =   26
         Top             =   5040
         Width           =   2055
      End
      Begin VB.CommandButton cmdInstallService 
         Caption         =   "&Install Service"
         Height          =   495
         Left            =   2400
         TabIndex        =   25
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lblServiceStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Service status:"
         Height          =   495
         Left            =   2040
         TabIndex        =   27
         Tag             =   "Service status: "
         Top             =   4200
         Width           =   5055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmControlPanel.frx":75E3
         Height          =   1455
         Left            =   840
         TabIndex        =   24
         Top             =   1320
         Width           =   7215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ViStart Service settings"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   23
         Top             =   480
         Width           =   7215
      End
   End
   Begin ViStart.NavigationBar naviBar 
      Height          =   8175
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3375
      _extentx        =   5953
      _extenty        =   14420
   End
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event onChangeSkin(skinName As CollectionItem)
Public Event onChangeOrb(szNewOrb As String)
Public Event onChangeRollover(szNewRollover As String)
Public Event onNavigationPanelChange()
Public Event onRequestAddMetroShortcut()

Private m_skinDir As String
Private m_orbDir As String
Private m_rolloverDir As String
'Private m_serviceInstalled As Boolean

Private m_navigationPane As ViNavigationPane
'Private m_serviceState As SERVICE_STATE

Private m_childSkinNameStrings As Collection

Private m_ElementY As Long

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1

Private m_logger As SeverityLogger

Property Get Logger() As SeverityLogger
    Set Logger = m_logger
End Property

Public Property Let NavigationPanel(ByRef newNavigationPane As ViNavigationPane)
    Set m_navigationPane = newNavigationPane
    InititalizeConfigureFrame
End Property

Sub ValidateSkin(ByVal szNewSkin As String)

        cmbChildThemes.Visible = False
        
    If FileCheck(m_skinDir & szNewSkin & "\") Then
        Set m_childSkinNameStrings = OptionsHelper.GetChildSkins(m_skinDir & szNewSkin & "\layout.xml")
        PopulateChildSkins

                If Settings.CurrentChildSkin <> vbNullString Then
                        cmbChildThemes.Visible = True
                'Else
                        'cmbChildThemes.Visible = False
                End If
    
        If Settings.CurrentSkin = vbNullString Or FileCheck(m_skinDir & Settings.CurrentSkin & "\") = False Then
            Settings.CurrentSkin = szNewSkin
        Else
            Dim newSkinCollectionItem As CollectionItem
            Set newSkinCollectionItem = New CollectionItem
            newSkinCollectionItem.Value = szNewSkin
                        
            RaiseEvent onChangeSkin(newSkinCollectionItem)
        End If
    Else
        MsgBox "This skin is broken!", vbCritical
    End If
    
End Sub

Private Sub chkDisableBottomLeftCorner_Click()
    MetroUtility.WindowsStartCorner_Disabled = CheckBoxToBoolean(chkDisableBottomLeftCorner.Value)

    Ensure_NukeMetro
End Sub

Private Sub chkDisableCharmsBar_Click()
    MetroUtility.CharmsBarBottom_Disabled = CheckBoxToBoolean(chkDisableCharmsBar.Value)
    MetroUtility.CharmsBarTop_Disabled = CheckBoxToBoolean(chkDisableCharmsBar.Value)

    Ensure_NukeMetro
End Sub

Private Sub chkDisableDragToClose_Click()
    MetroUtility.DragToClose_Disabled = CheckBoxToBoolean(chkDisableDragToClose.Value)

    Ensure_NukeMetro
End Sub

Private Sub chkHotCorners_Click()
    If chkHotCorners.Value = vbChecked Then
        chkDisableCharmsBar.Value = vbChecked
        chkDisableDragToClose.Value = vbChecked
        chkDisableBottomLeftCorner.Value = vbChecked
    Else
        chkDisableCharmsBar.Value = vbUnchecked
        chkDisableDragToClose.Value = vbUnchecked
        chkDisableBottomLeftCorner.Value = vbUnchecked
    End If
End Sub

Private Sub chkProgramMenuFirst_Click()
    Settings.ShowProgramsFirst = CheckBoxToBoolean(chkProgramMenuFirst.Value)
End Sub

Private Sub chkSkipMetroScreen_Click()
    MetroUtility.SkipMetro_Enabled = CheckBoxToBoolean(chkSkipMetroScreen.Value)
    
    Ensure_NukeMetro
End Sub

Private Sub chkStartWithWindows_Click()

Dim runAtStartupRegKey As RegistryKey
    Set runAtStartupRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run")

    If chkStartWithWindows.Value = vbChecked Then
        runAtStartupRegKey.SetValue "ViStart", AppPath & App.EXEName & ".exe"
    Else
        runAtStartupRegKey.DeleteValue "ViStart"
    End If
End Sub

Private Sub chkShowSplashScreen_Click()
    Settings.ShowSplashScreen = CheckBoxToBoolean(chkShowSplashScreen.Value)
End Sub

Private Sub Ensure_NukeMetro()

Dim runAtStartupRegKey As RegistryKey
    Set runAtStartupRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run")
    
    If runAtStartupRegKey Is Nothing Then
        Logger.Error "Unable to open registry key", "Ensure_NukeMetro"
        Exit Sub
    End If

    If chkDisableCharmsBar.Value = vbChecked Or _
       chkDisableDragToClose.Value = vbChecked Or _
       chkDisableBottomLeftCorner.Value = vbChecked Or _
       chkSkipMetroScreen.Value = vbChecked Then
        
        runAtStartupRegKey.SetValue "NukeMetro", """" & AppPath & App.EXEName & ".exe" & """" & " " & _
            "/nuke_metro"
    Else
        runAtStartupRegKey.DeleteValue "NukeMetro"
    End If
End Sub

Private Sub chkSystemTray_Click()
    Settings.EnableTrayIcon = CheckBoxToBoolean(chkSystemTray.Value)
      
    If Settings.EnableTrayIcon Then
        frmEvents.AddIconToTray
    Else
        frmEvents.DeleteIconFromTray
    End If
End Sub

Private Sub chkUserPicture_Click()
    Settings.ShowUserPicture = CheckBoxToBoolean(chkUserPicture.Value)
End Sub

Private Sub cmbChildThemes_Click()
    Dim newSkinCollectionItem As CollectionItem
    Set newSkinCollectionItem = New CollectionItem
    
    newSkinCollectionItem.Key = m_childSkinNameStrings(cmbChildThemes.listIndex + 1).Key
    newSkinCollectionItem.Value = cmbThemes.text
    
    RaiseEvent onChangeSkin(newSkinCollectionItem)
End Sub

Private Sub cmbStartOrbs_Change()
    cmbStartOrbs_Click
End Sub

Private Sub cmbStartOrbs_Click()
    
    If cmbStartOrbs.listIndex = 0 Then
        RaiseEvent onChangeOrb(vbNullString)
    Else
        RaiseEvent onChangeOrb(cmbStartOrbs.text)
    End If
End Sub

Private Sub cmbRollover_Change()
    cmbRollover_Click
End Sub

Private Sub cmbRollover_Click()
    
    If cmbRollover.listIndex = 0 Then
                Settings.CurrentRollover = vbNullString
                g_rolloverPath = sCon_AppDataPath & "_skins\" & Settings.CurrentSkin & "\rollover\"

                RaiseEvent onChangeRollover(vbNullString)
    Else
                Settings.CurrentRollover = cmbRollover.text
        g_rolloverPath = sCon_AppDataPath & "_rollover\" & Settings.CurrentRollover & "\"

        RaiseEvent onChangeRollover(cmbRollover.text)
    End If

End Sub

Private Sub cmbThemes_Click()
    ValidateSkin cmbThemes.text
End Sub

Private Sub cmbWindowsKey_Change()
    cmbWindowsKey_Click
End Sub

Private Sub cmbWindowsKey_Click()
        
    If cmbWindowsKey.listIndex = 0 Then
        Settings.CatchLeftWindowsKey = True
        Settings.CatchRightWindowsKey = True
    ElseIf cmbWindowsKey.listIndex = 1 Then
        Settings.CatchLeftWindowsKey = True
        Settings.CatchRightWindowsKey = False
    ElseIf cmbWindowsKey.listIndex = 2 Then
        Settings.CatchLeftWindowsKey = False
        Settings.CatchRightWindowsKey = True
    ElseIf cmbWindowsKey.listIndex = 3 Then
        Settings.CatchLeftWindowsKey = False
        Settings.CatchRightWindowsKey = False
    End If
    
End Sub

Private Sub cmbWindowsOrb_Change()
    cmbWindowsOrb_Click
End Sub

Private Sub cmbWindowsOrb_Click()
    If cmbWindowsOrb.listIndex = 0 Then
        Settings.StartButtonShowsWindowsMenu = False
    Else
        Settings.StartButtonShowsWindowsMenu = True
    End If
End Sub

Private Sub cmdInstallTheme_Click()

Dim szNewThemeFile As String
Dim szNewThemeName As String

    szNewThemeFile = BrowseForFile(vbNullString, "ViStart Theme (*.vistart-theme);*.vistart-theme", GetPublicString("strViStartTheme"), Me.hWnd)
    
    If FileExists(szNewThemeFile) Then
        If Not InstallTheme(szNewThemeFile, szNewThemeName) Then
            MsgBox "Error installing new theme!", vbCritical
            Exit Sub
        End If
        
        ListSkins
        SelectSkinByName szNewThemeName
    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdMoreOrbs_Click()
    AppLauncherHelper.ShellEx "http://lee-soft.com/skins/"
End Sub

Private Sub cmdMoreThemes_Click()
    AppLauncherHelper.ShellEx "http://lee-soft.com/skins/"
End Sub

Private Sub cmdPickImage_Click()
    frmStartOrb.PromptUserSelectNewOrb
    ListOrbs
End Sub

Private Sub Command1_Click()
    RaiseEvent onRequestAddMetroShortcut
End Sub

Private Sub cmdRestoreSpecialFolders_Click()

    If Not SpecialFoldersHelper.RestoreDefaultFolders = Success Then
        MsgBox "Unable to restore default folders to explorer.", vbCritical
        Exit Sub
    End If
    
    MsgBox "Succesfully restored default folders!", vbInformation
End Sub

Private Sub cmdShowMetroShortcut_Click()
    RaiseEvent onRequestAddMetroShortcut
    CheckMetroShortcut
    
End Sub

Private Sub Form_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)
    Set m_childSkinNameStrings = New Collection
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, 1, True

    Me.Height = (picFrame(0).Height + 28) * Screen.TwipsPerPixelY
    Me.Width = (naviBar.Width + picFrame(0).Width) * Screen.TwipsPerPixelX
    
    'InitializeStyleFrame
    'InitializeDesktopFrame

    naviBar.AddItem GetPublicString("strStyle")
    naviBar.AddItem GetPublicString("strConfigure")
    naviBar.AddItem GetPublicString("strDesktop")
    naviBar.AddItem GetPublicString("strAbout")
    naviBar.AddItem "Donate"
    
    naviBar.SelectedIndex = 1
    
    ' Set up scroll bars:
    Set m_cScroll = New cScrollBars
    m_cScroll.Create picFrame(1).hWnd
    m_cScroll.SmallChange(efsVertical) = MenuItem(0).Height \ Screen.TwipsPerPixelY + 2
    
    InitializeFrame "about.vml"

    frmControlPanel.Caption = GetPublicString("strViStartControlPanel")

    Label1.Caption = GetPublicString("strWhichStartMenu")
    cmdInstallTheme.Caption = GetPublicString("strInstall")
        
    Label2.Caption = GetPublicString("strWhatStarOrb")
    cmdPickImage.Caption = GetPublicString("strPick")
 
    Label7.Caption = GetPublicString("strWhatRollover")
 
    Label3.Caption = GetPublicString("strWhatToSee")
    Label4.Caption = GetPublicString("strWhatToSeeOnRight")
        
    chkProgramMenuFirst.Caption = GetPublicString("strProgramsFirst")
    chkUserPicture.Caption = GetPublicString("strShowUserPicture")
        
    Label6.Caption = GetPublicString("strDesktopSettings")
        
    chkSystemTray.Caption = GetPublicString("strShowTrayIcon")
    chkStartWithWindows.Caption = GetPublicString("strStartWithWindows")
    chkShowSplashScreen.Caption = GetPublicString("strSplash")
        
    cmdShowMetroShortcut.Caption = GetPublicString("strRestoreStartMenu")
        
    Label5.Caption = GetPublicString("strW8Features")
        
    chkHotCorners.Caption = GetPublicString("strHotCorners")
    chkDisableCharmsBar.Caption = GetPublicString("strDisableCharmsBar")
    chkDisableDragToClose.Caption = GetPublicString("strDisableDragToClose")
    chkDisableBottomLeftCorner.Caption = GetPublicString("strDisableBottomLeftCorner")
    chkSkipMetroScreen.Caption = GetPublicString("strSkipMetroScreen")
        
    Label11.Caption = GetPublicString("strW8FeaturesWarning")

    'lblSubText.Caption = GetPublicString("strCopyright")
        
End Sub

Sub NavigateToPanel(ByVal szPanelName As String)
    naviBar.NavigateToItem szPanelName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not g_Exiting Then
        Cancel = True
        Me.Hide
        Exit Sub
    Else
        Set m_cScroll = Nothing
    End If

End Sub

Private Sub MenuItem_onChanged(Index As Integer)
    RaiseEvent onNavigationPanelChange
End Sub

Private Sub naviBar_onClick(theIndex As Long)
    On Error GoTo Handler
    
    If theIndex = 1 Then
        InititalizeConfigureFrame
    ElseIf theIndex = 2 Then
        'CheckMetro
        CheckMetroShortcut
    ElseIf theIndex = 3 Then
        InitializeFrame "about.vml"
        timReloadAbout.Enabled = True
    ElseIf theIndex = 4 Then
        theIndex = 3
        InitializeFrame "info.vml"
        timReloadAbout.Enabled = True
    End If

    picFrame(theIndex).ZOrder 0
    picFrame(theIndex).Move picFrame(0).Left, picFrame(0).Top
    
    picFrame(theIndex).Visible = True

    Exit Sub
Handler:
End Sub

Sub CheckMetroShortcut()

Dim buttonEnabled As Boolean

    If g_Windows8 Or g_Windows81 Then
        If Settings.Programs.ExistsInPinned("!default_menu") = False Or Settings.Programs.ExistsInPinned("explorer shell:::{2559a1f8-21d7-11d4-bdaf-00c04f60b9f0}") = False Then
            cmdShowMetroShortcut.Enabled = True
        Else
            cmdShowMetroShortcut.Enabled = False
        End If
    Else
        If Settings.Programs.ExistsInPinned("!default_menu") = False Then
            cmdShowMetroShortcut.Enabled = True
        Else
            cmdShowMetroShortcut.Enabled = False
        End If
    End If
End Sub

Sub InitializeFrame(ByVal frameName As String)
    On Error GoTo Finally

Dim infoXMLText As String
    infoXMLText = LoadStringFromResource(frameName, "VML")
    
    RenderXMLFrame infoXMLText
    
Finally:
    RenderXMLFrame infoXMLText
End Sub

Sub InitializeAboutSkinFrame()

Dim infoTextStream As StreamReader
Dim infoXMLText As String

    Set infoTextStream = New StreamReader
    
    On Error GoTo Finally
    
    infoTextStream.OpenStream g_resourcesPath & "\info.xml"
    infoXMLText = infoTextStream.ReadToEnd
    
Finally:
    RenderXMLFrame infoXMLText
End Sub

Sub RenderXMLFrame(ByVal sourceXML As String)
    VMLViewer.XML = sourceXML
End Sub

Sub InitializeDesktopFrame()
    chkSystemTray.Value = BooleanToCheckBox(Settings.EnableTrayIcon)
    chkShowSplashScreen.Value = BooleanToCheckBox(Settings.ShowSplashScreen)
    
    cmbWindowsKey.Clear
    cmbWindowsKey.AddItem GetPublicString("strBothWinKeysViStart")
    cmbWindowsKey.AddItem GetPublicString("strLeftWinKey")
    cmbWindowsKey.AddItem GetPublicString("strRightWinKey")
    cmbWindowsKey.AddItem GetPublicString("strBothWinKeys")
    
    cmbWindowsOrb.Clear
    cmbWindowsOrb.AddItem GetPublicString("strStartViStart")
    cmbWindowsOrb.AddItem GetPublicString("strStartWinMenu")
    
    If Settings.CatchLeftWindowsKey And Settings.CatchRightWindowsKey Then
        cmbWindowsKey.listIndex = 0
    End If
        
    If Settings.CatchLeftWindowsKey = True And Settings.CatchRightWindowsKey = False Then
        cmbWindowsKey.listIndex = 1
    End If
        
    If Settings.CatchRightWindowsKey = True And Settings.CatchLeftWindowsKey = False Then
        cmbWindowsKey.listIndex = 2
    End If
    
    If Settings.CatchLeftWindowsKey = False And Settings.CatchRightWindowsKey = False Then
        cmbWindowsKey.listIndex = 3
    End If
    
    If Settings.StartButtonShowsWindowsMenu Then
        cmbWindowsOrb.listIndex = 1
    Else
        cmbWindowsOrb.listIndex = 0
    End If
    
    If g_Windows8 Or g_Windows81 Then
        picWindows8.Visible = True
        cmdShowMetroShortcut.Caption = "Restore Windows Metro Shortcuts"
    End If
    
    
    If MetroUtility.CharmsBarBottom_Disabled Or MetroUtility.CharmsBarTop_Disabled Then
        chkDisableCharmsBar.Value = vbChecked
    End If
    
    chkDisableDragToClose.Value = BooleanToCheckBox(MetroUtility.DragToClose_Disabled)
    chkDisableBottomLeftCorner.Value = BooleanToCheckBox(MetroUtility.WindowsStartCorner_Disabled)
    chkSkipMetroScreen.Value = BooleanToCheckBox(MetroUtility.SkipMetro_Enabled)
    chkStartWithWindows.Value = BooleanToCheckBox(StartsWithWindows)
    
    If chkDisableCharmsBar.Value = vbChecked And chkDisableBottomLeftCorner.Value = vbChecked And chkDisableDragToClose.Value = vbChecked Then
        chkHotCorners.Value = vbChecked
    End If
End Sub

Sub InitializeStyleFrame()
    m_skinDir = sCon_AppDataPath & "_skins\"
    m_orbDir = sCon_AppDataPath & "_orbs\"
    m_rolloverDir = sCon_AppDataPath & "_rollover\"
    
    ListSkins
    ListOrbs
    ListRollovers
        
        ' Hide rollover option when no _rollover folder exist
        If cmbRollover.ListCount > 1 Then
                Label7.Visible = True
                cmbRollover.Visible = True
        Else
                Label7.Visible = False
                cmbRollover.Visible = False
        End If
        
End Sub

Sub InititalizeConfigureFrame()
    If m_navigationPane Is Nothing Then Exit Sub

Dim thisOption As Object

    If ClearNavigationItems = False Then
        Exit Sub
    End If
    
    For Each thisOption In m_navigationPane.NavigationOptions
        AddNavigationItem thisOption
    Next
    
    chkUserPicture.Value = BooleanToCheckBox(Settings.ShowUserPicture)
    chkProgramMenuFirst.Value = BooleanToCheckBox(Settings.ShowProgramsFirst)
    
    
End Sub

Function ClearNavigationItems() As Boolean
    On Error GoTo Handler

    If MenuItem.count = 1 Then
        ClearNavigationItems = True
        Exit Function
    End If
    
Dim labelIndex As Long

    For labelIndex = 1 To MenuItem.UBound
        Unload Me.MenuItem(labelIndex)
        picClient_1.Height = picClient_1.Height - (MenuItem(0).Height + 2)
    Next
    
    ClearNavigationItems = True
    Exit Function
Handler:
    If Err.Number = 365 Then
        timDelayUnload.Enabled = True
        Err.Clear
    Else
        MsgBox Err.Number, vbCritical
    End If

End Function

Sub AddNavigationItem(ByRef objNavigationItem As Object)
    
    Load MenuItem(MenuItem.count)
    With MenuItem(MenuItem.UBound)
        .Left = MenuItem(0).Left
        .Top = MenuItem(.Index - 1).Top + MenuItem(0).Height + 2

        .Source = objNavigationItem
        .Visible = True
    End With
    
    picClient_1.Height = picClient_1.Height + (MenuItem(0).Height + 2)
End Sub

Sub ListOrbs()
    On Error GoTo Handler

    cmbStartOrbs.Clear
    cmbStartOrbs.AddItem GetPublicString("strSkinDefaultOrb")

'Dim thisSubFolder As Scripting.Folder
Dim thisFolder As Scripting.Folder
Dim thisOrb As Scripting.File
Dim foundOrb As Boolean

    If FSO.FolderExists(m_orbDir) = False Then
        Exit Sub
    End If

    Set thisFolder = FSO.GetFolder(m_orbDir)
    
    For Each thisOrb In thisFolder.Files

        cmbStartOrbs.AddItem thisOrb.name
        
        If Not foundOrb Then
            If LCase$(thisOrb.name) = LCase$(Settings.CurrentOrb) Then
                cmbStartOrbs.listIndex = cmbStartOrbs.ListCount - 1
                foundOrb = True
            End If
        End If
    Next
    
    If Not foundOrb Then
        cmbStartOrbs.listIndex = 0
    End If

    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Sub ListRollovers()

    On Error GoTo Handler

    cmbRollover.Clear
    cmbRollover.AddItem GetPublicString("strSkinDefaultRollover")
        
Dim thisSubFolder As Scripting.Folder
Dim thisFolder As Scripting.Folder

    If FSO.FolderExists(m_rolloverDir) = False Then
        Exit Sub
    End If

    Set thisFolder = FSO.GetFolder(m_rolloverDir)
    
    For Each thisSubFolder In thisFolder.SubFolders
        
        If FSO.FileExists(m_rolloverDir & thisSubFolder.name & "\computer.png") Then
            cmbRollover.AddItem thisSubFolder.name

            If LCase$(thisSubFolder.name) = LCase$(Settings.CurrentRollover) Then
                cmbRollover.listIndex = cmbRollover.ListCount - 1
            End If

        End If
    Next
        
        If Settings.CurrentRollover = vbNullString Then
                cmbRollover.listIndex = 0
        End If
        
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Sub SelectSkinByName(szSkinName As String)

Dim listIndex As Long

    For listIndex = 1 To cmbThemes.ListCount
        If LCase$(cmbThemes.List(listIndex)) = LCase$(szSkinName) Then
            cmbThemes.listIndex = listIndex
        End If
    Next

End Sub

Sub PopulateChildSkins()

    cmbChildThemes.Clear

Dim sourceItem As CollectionItem
    For Each sourceItem In m_childSkinNameStrings
        cmbChildThemes.AddItem sourceItem.Value
    Next

    If cmbChildThemes.ListCount > 0 Then
        cmbChildThemes.listIndex = 0
    End If
End Sub

Sub ListSkins()
    On Error GoTo Handler

    cmbThemes.Clear

Dim thisSubFolder As Scripting.Folder
Dim thisFolder As Scripting.Folder

    If FSO.FolderExists(m_skinDir) = False Then
        Exit Sub
    End If

    Set thisFolder = FSO.GetFolder(m_skinDir)
    
    For Each thisSubFolder In thisFolder.SubFolders
        If FileCheck(m_skinDir & thisSubFolder.name & "\") Then
            cmbThemes.AddItem thisSubFolder.name
            
            If LCase$(thisSubFolder.name) = LCase$(Settings.CurrentSkin) Then
                cmbThemes.listIndex = cmbThemes.ListCount - 1
            End If
        End If
    Next

    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub picClient_1_Resize()
On Error Resume Next

       Dim lHeight As Long
       Dim lWidth As Long
       Dim lProportion As Long
          
       ' Pixels are the minimum change size for a screen object.
       ' Therefore we set the scroll bars in pixels.
    
       lHeight = (picClient_1.ScaleHeight - picFrame(1).ScaleHeight) \ Screen.TwipsPerPixelY
       If (lHeight > 0) Then
          lProportion = lHeight \ (picClient_1.ScaleHeight \ Screen.TwipsPerPixelY) + 1
          m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
          m_cScroll.Max(efsVertical) = lHeight
          m_cScroll.Visible(efsVertical) = True
       Else
          m_cScroll.Visible(efsVertical) = False
       End If
       
End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
On Error Resume Next

   If (m_cScroll.Visible(eBar)) Then
      If (eBar = efsVertical) Then
         picClient_1.Top = -m_cScroll.Value(eBar) * Screen.TwipsPerPixelY
      End If
   Else
      picClient_1.Move 0, 0
   End If
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   m_cScroll_Change eBar
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub timDelayUnload_Timer()
    timDelayUnload.Enabled = False
    
    InititalizeConfigureFrame
End Sub

Private Sub timReloadAbout_Timer()
    'InitializeFrame "about.vml"
    timReloadAbout.Enabled = False
End Sub

Private Sub VMLViewer_LinkClicked(ByVal tag As String)

    If tag = "_ABOUT_SKIN" Then
        InitializeAboutSkinFrame
    ElseIf tag <> vbNullString Then
        AppLauncherHelper.ShellEx tag
    End If

End Sub
