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
      Index           =   3
      Left            =   5120
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   577
      TabIndex        =   21
      Top             =   940
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "<New Text Object>"
         Height          =   375
         Index           =   0
         Left            =   1080
         MouseIcon       =   "frmControlPanel.frx":74F2
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   1560
         Visible         =   0   'False
         Width           =   6660
      End
      Begin VB.Label lblAurthor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<PLACE HOLDER>"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   380
         Left            =   1080
         MouseIcon       =   "frmControlPanel.frx":7644
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Tag             =   "http://lee-soft.com"
         Top             =   7420
         Visible         =   0   'False
         Width           =   6660
      End
      Begin VB.Label lblViStartTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ViStart 8.1"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   22
         Top             =   420
         Width           =   7455
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
            TabIndex        =   47
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
            _ExtentX        =   11880
            _ExtentY        =   1085
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmControlPanel.frx":7796
            Height          =   555
            Left            =   150
            TabIndex        =   46
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
            TabIndex        =   44
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
         TabIndex        =   41
         Top             =   3410
         Width           =   5055
      End
      Begin VB.CommandButton cmdShowMetroShortcut 
         Caption         =   "Restore Windows Start Menu Shortcut"
         Height          =   400
         Left            =   720
         TabIndex        =   38
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
         TabIndex        =   30
         Top             =   4320
         Visible         =   0   'False
         Width           =   8055
         Begin VB.CheckBox chkSkipMetroScreen 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Automatically go to desktop when I log in "
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   2880
            Width           =   5055
         End
         Begin VB.CheckBox chkDisableBottomLeftCorner 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable bottom left (Start) hot corner"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   2400
            Width           =   5055
         End
         Begin VB.CheckBox chkDisableDragToClose 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Drag to close"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   1920
            Width           =   5055
         End
         Begin VB.CheckBox chkDisableCharmsBar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable CharmsBar"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1440
            Width           =   5055
         End
         Begin VB.CheckBox chkHotCorners 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable all Windows 8 hot corners"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   960
            Width           =   5055
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Windows 8 related features require a restart to take effect "
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   840
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   360
            Width           =   7695
         End
      End
      Begin VB.CheckBox chkStartWithWindows 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start ViStart with Windows"
         Height          =   375
         Left            =   720
         TabIndex        =   29
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
         ItemData        =   "frmControlPanel.frx":7847
         Left            =   720
         List            =   "frmControlPanel.frx":784E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   4935
      End
      Begin VB.ComboBox cmbWindowsKey 
         Height          =   360
         ItemData        =   "frmControlPanel.frx":7861
         Left            =   720
         List            =   "frmControlPanel.frx":7868
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
         ItemData        =   "frmControlPanel.frx":787B
         Left            =   720
         List            =   "frmControlPanel.frx":787D
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1800
         Width           =   4935
      End
      Begin VB.ComboBox cmbRollover 
         Height          =   360
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   43
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
         ItemData        =   "frmControlPanel.frx":787F
         Left            =   720
         List            =   "frmControlPanel.frx":7881
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3360
         Width           =   4935
      End
      Begin VB.ComboBox cmbThemes 
         Height          =   360
         ItemData        =   "frmControlPanel.frx":7883
         Left            =   720
         List            =   "frmControlPanel.frx":7885
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
         TabIndex        =   42
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
      TabIndex        =   23
      Top             =   7340
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdStartService 
         Caption         =   "&Start Service"
         Height          =   495
         Left            =   4800
         TabIndex        =   27
         Top             =   5040
         Width           =   2055
      End
      Begin VB.CommandButton cmdInstallService 
         Caption         =   "&Install Service"
         Height          =   495
         Left            =   2400
         TabIndex        =   26
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lblServiceStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Service status:"
         Height          =   495
         Left            =   2040
         TabIndex        =   28
         Tag             =   "Service status: "
         Top             =   4200
         Width           =   5055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmControlPanel.frx":7887
         Height          =   1455
         Left            =   840
         TabIndex        =   25
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
         TabIndex        =   24
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
      _ExtentX        =   5953
      _ExtentY        =   14420
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

    On Error GoTo Handler
    SpecialFoldersHelper.RestoreDefaultFolders
    MsgBox "Succesfully restored the explorer default folders", vbInformation
    
    Exit Sub
Handler:
    Logger.Error Err.Description, "cmdRestoreSpecialFolders_Click"
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
    SetIcon Me.hWnd, "APPICON", True

    Me.Height = (picFrame(0).Height + 28) * Screen.TwipsPerPixelY
    Me.Width = (naviBar.Width + picFrame(0).Width) * Screen.TwipsPerPixelX
    
    InitializeStyleFrame
    InitializeDesktopFrame

    naviBar.AddItem GetPublicString("strStyle")
    naviBar.AddItem GetPublicString("strConfigure")
    naviBar.AddItem GetPublicString("strDesktop")
    naviBar.AddItem GetPublicString("strAbout")
    
    naviBar.SelectedIndex = 1
    
    ' Set up scroll bars:
    Set m_cScroll = New cScrollBars
    m_cScroll.Create picFrame(1).hWnd
    m_cScroll.SmallChange(efsVertical) = MenuItem(0).Height \ Screen.TwipsPerPixelY + 2
    
    InitializeAboutFrame

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

Private Sub lblAurthor_Click()
    AppLauncherHelper.ShellEx lblAurthor.Tag
End Sub

Private Sub lblText_Click(Index As Integer)
    If lblText(Index).Tag = "_ABOUT_SKIN" Then
        InitializeAboutSkinFrame
    ElseIf lblText(Index).Tag <> vbNullString Then
        AppLauncherHelper.ShellEx lblText(Index).Tag
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

'Sub InitializeServiceFrame()

    'CheckService

'End Sub

'Private Sub CheckService()

     'If GetServiceConfig() = 0 Then
        'm_serviceInstalled = True
        'm_serviceState = GetServiceStatus()
        
        'cmdInstallService.Caption = "&Remove Service"
        
        'Select Case m_serviceState
        
            'Case SERVICE_RUNNING
            
                'SetServiceStatus "Running", vbGreen
            
                'cmdInstallService.Enabled = False
                'cmdStartService.Caption = "&Stop Service"
                'cmdStartService.Enabled = True
                
            'Case SERVICE_STOPPED
            
                'SetServiceStatus "Stopped"
            
                'cmdInstallService.Enabled = True
                'cmdStartService.Caption = "&Start Service"
                'cmdStartService.Enabled = True
                
            'Case Else
            
                'SetServiceStatus "Unknown..", vbRed
            
                'cmdInstallService.Enabled = False
                'cmdStartService.Enabled = False
                
        'End Select
    'Else
    
        'SetServiceStatus "Not Installed"
    
        'cmdInstallService.Caption = "&Install Service"
    
        'm_serviceInstalled = False
    
        'cmdStartService.Enabled = False
        'cmdInstallService.Enabled = True
    'End If

'End Sub

'Private Sub SetServiceStatus(ByVal szNewStatus As String, Optional ByVal newForeColour As ColorConstants = vbBlack)

    'lblServiceStatus.Caption = lblServiceStatus.Tag & szNewStatus
    'lblServiceStatus.ForeColor = newForeColour

'End Sub

'Private Sub cmdStartService_Click()

    'CheckService
    'If Not cmdStartService.Enabled Then Exit Sub
    'cmdStartService.Enabled = False
    
    'If m_serviceState = SERVICE_RUNNING Then
        'NTServiceControl.StopNTService
    'ElseIf m_serviceState = SERVICE_STOPPED Then
        'NTServiceControl.StartNTService
    'End If
    
    'CheckService

'End Sub

'Private Sub cmdInstallService_Click()
    'CheckService
    
    'If Not cmdInstallService.Enabled Then Exit Sub
    'cmdInstallService.Enabled = False
    
    'If m_serviceInstalled Then
        'NTServiceControl.DeleteNTService
    'Else
        'NTServiceControl.SetNTService
    'End If
    
    'CheckService
'End Sub

Sub InitializeAboutFrame()

Dim infoXMLText As String
    infoXMLText = LoadStringFromResource("about.xml", "SCHEMA")
    
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

    lblViStartTitle.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    ClearAboutText

Dim xmlDoc As DOMDocument
Dim skinInfoXML As IXMLDOMElement
Dim thisChild As IXMLDOMElement
Dim thisObject As Object

    Set xmlDoc = New DOMDocument
    
    If xmlDoc.loadXML(sourceXML) = False Then
        Exit Sub
    End If
 
    m_ElementY = lblViStartTitle.Top + lblViStartTitle.Height + 10
    
    Set skinInfoXML = xmlDoc.firstChild
    For Each thisObject In skinInfoXML.childNodes
        If TypeName(thisObject) = "IXMLDOMElement" Then
            Set thisChild = thisObject
    
            Select Case LCase$(thisChild.tagName)
            
            Case "a"
                ParseHref thisChild
            
            Case "p"
                ParseParagraph thisChild
                
            'Case "description"
                'lblDescription.Caption = thisChild.Text
            
            End Select
        End If
    Next

End Sub

Private Sub ClearAboutText()
    On Error GoTo Handler
    
    While lblText.count > 1
        Unload lblText(lblText.UBound)
    Wend
    
Handler:
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

Private Sub ParseHref(ByRef theText As IXMLDOMElement)
    On Error GoTo Handler

Dim theCaption As String
Dim theHref As String
Dim theAlignment As AlignmentConstants
Dim theFontStyle As Long
Dim realPosition_X As Long

Dim theDimensions As RECTF
Dim theLabelIndex As Long

    theAlignment = vbLeftJustify
    theCaption = theText.text
    
    If Not IsNull(theText.getAttribute("href")) Then
        theHref = theText.getAttribute("href")
    End If
    
    If Not IsNull(theText.getAttribute("align")) Then
        Select Case LCase(theText.getAttribute("align"))
        
        Case "left"
            theAlignment = vbLeftJustify
        
        Case "right"
            theAlignment = vbRightJustify
            
        Case "center"
            theAlignment = vbCenter
            
        End Select
    End If
    
    With AddText(theCaption, FontStyleRegular, theAlignment)
        .Tag = theHref
        .FontUnderline = True
        .ForeColor = vbBlue
        .MousePointer = 99
    End With
    
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub ParseParagraph(ByRef theText As IXMLDOMElement)

Dim theCaption As String
Dim theFontStyleText As String
Dim realPosition_X As Long

Dim theDimensions As RECTF

Dim theLabelIndex As Long
Dim theFontStyle As fontStyle
Dim theAlignment As AlignmentConstants

    theAlignment = vbLeftJustify

    If Not IsNull(theText.text) Then
        theCaption = theText.text
    End If
    
    If Not IsNull(theText.getAttribute("align")) Then
        Select Case LCase(theText.getAttribute("align"))
        
        Case "left"
            theAlignment = vbLeftJustify
        
        Case "right"
            theAlignment = vbRightJustify
            
        Case "center"
            theAlignment = vbCenter
            
        End Select
    End If
    
    If Not IsNull(theText.getAttribute("style")) Then
        theFontStyleText = theText.getAttribute("style")
    End If
    
    theFontStyle = FontStyleRegular
    
    Select Case LCase$(theFontStyleText)
    
    Case "bold"
        theFontStyle = FontStyleBold
        
    Case "italic"
        theFontStyle = FontStyleItalic
        
    Case "bold|italic"
        theFontStyle = FontStyleBoldItalic
        
    Case "underline"
        theFontStyle = FontStyleUnderline
        
    Case "strikeout"
        theFontStyle = FontStyleStrikeout
    
    End Select
    
    AddText theCaption, theFontStyle, theAlignment
End Sub

Private Function AddText(ByVal theCaption As String, Optional ByVal theFontStyle As fontStyle = FontStyleRegular, Optional TextAlignment As AlignmentConstants = vbLeftJustify) As Label
    'If theCaption = vbNullString Then Exit Function
    On Error GoTo Handler

Dim newHeight As Long

    theCaption = Replace(theCaption, "%skin%", Settings.CurrentSkin)

    Load lblText(lblText.count)

    With lblText(lblText.UBound)
        .ForeColor = vbBlack
        
        If theFontStyle = FontStyleBold Then
            .FontBold = True
        ElseIf theFontStyle = FontStyleBoldItalic Then
            .FontBold = True
            .FontItalic = True
        ElseIf theFontStyle = FontStyleItalic Then
            .FontItalic = True
        ElseIf theFontStyle = FontStyleStrikeout Then
            .FontStrikethru = True
        ElseIf theFontStyle = FontStyleUnderline Then
            .FontUnderline = True
        End If
        
        .Caption = theCaption
        .AutoSize = True
        .MousePointer = 0

        'If .Width > lblAurthor.Width Then
        '    newHeight = (Ceiling(.Width / lblAurthor.Width)) * (lblAurthor.Height)
        '    .AutoSize = False
       '
       '     .Height = newHeight
       '     .Width = lblAurthor.Width
       '     .Left = lblAurthor.Left
       ' End If

        .Visible = True
        .Width = lblAurthor.Width

        .Top = m_ElementY
        m_ElementY = m_ElementY + .Height
        
        .Caption = .Caption
        .Alignment = TextAlignment
    End With
    
    'm_rectPosition.Top = m_Y
    'm_rectPosition.Left = realPosition_X
    'm_rectPosition.Width = m_rectPosition.Left + theDimensions.Width
    'm_rectPosition.Height = m_rectPosition.Top + theDimensions.Height
    'm_path.AddString theCaption, m_fontF, theFontStyle, CSng(theFontSize), m_rectPosition, 0
    
    Set AddText = lblText(lblText.UBound)
    
    Exit Function
Handler:
    MsgBox Err.Description, vbCritical
End Function

Private Sub timReloadAbout_Timer()
    InitializeAboutFrame
    timReloadAbout.Enabled = False
End Sub
