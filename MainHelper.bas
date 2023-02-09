Attribute VB_Name = "MainHelper"
Option Explicit

Public Layout As LayoutParser
Public Registry As clsShellReg
Public g_winLoading As frmAbout
Public Settings As ViSettings
Public CmdLine As CommandLine
Public MetroUtility As Windows8Utility
Public IconManager As CIconManager

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
   
Private Const ICC_USEREX_CLASSES = &H200

Private sOldIconSize As String

Private m_IgnorePreviousInstance As Boolean
Public g_Exiting As Boolean

Public g_rolloverImage As GDIPImage
Public g_viOrb_fullHeight As Boolean
Public g_WDSInitialized As Boolean
Public g_resourcesPath As String
Public g_rolloverPath As String
Public g_startButtonFindAttempts As Long

Private Const EXIT_PROGRAM As Long = 1

Public Function WaitForDesktop()

    ShellHelper.UpdateHwnds
    
    While IsWindow(ShellHelper.g_lnghwndTaskBar) = APIFALSE Or IsWindowVisible(ShellHelper.g_lnghwndTaskBar) = APIFALSE
        ShellHelper.UpdateHwnds
        Sleep 500
    Wend

End Function

Public Function GetResourceBitmap(ByVal theName) As IPictureDisp
    On Error GoTo Handler

    If FileExists(ResourcesPath & CStr(theName) & ".bmp") Then
        Set GetResourceBitmap = LoadPicture(ResourcesPath & CStr(theName) & ".bmp")
    Else
        Set GetResourceBitmap = LoadResPicture(theName, vbResBitmap)
    End If
Handler:
End Function

Public Function GetProgramMenuBackColour() As Long
    On Error GoTo Handler
    GetProgramMenuBackColour = vbWhite
    
Dim xmlLayout As New DOMDocument
Dim subElement As IXMLDOMElement

    If xmlLayout.Load(ResourcesPath & "layout.xml") = False Then
        GetProgramMenuBackColour = vbWhite
        Exit Function
    End If
    
    Set subElement = xmlLayout.selectSingleNode("startmenu_base//vielement[@id='programmenu']")
    GetProgramMenuBackColour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "backcolour", "#ffffff"))
    
    Exit Function
Handler:
    
End Function

Public Function FileCheck(szSourcePath As String) As Boolean

    If Not FileExists(szSourcePath & "startmenu.png") Or _
       Not FileExists(szSourcePath & "userframe.png") Or _
       Not FileExists(szSourcePath & "bottombuttons_arrow.png") Or _
       Not FileExists(szSourcePath & "bottombuttons_shutdown.png") Or _
       Not FileExists(szSourcePath & "allprograms.png") Or _
       Not FileExists(szSourcePath & "button.png") Or _
       Not FileExists(szSourcePath & "programs_arrow.png") Then

       FileCheck = False: Exit Function
    End If

    FileCheck = True

End Function

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Function InitClasses_IfNeeded() As Boolean
    If Not Registry Is Nothing Then
        InitClasses_IfNeeded = True
        Exit Function
    End If

    InitCommonControlsVB
    
    Set Registry = New clsShellReg
    Set Layout = New LayoutParser
    Set MetroUtility = New Windows8Utility
    Set IconManager = New CIconManager
    
    If Not Layout.ParseStartMenu(ResourcesPath & "layout.xml").ErrorCode = 0 Then
        Exit Function
    End If
    
    SetOpenSaveDocs
    
    g_viOrb_fullHeight = Layout.ViOrb_FullHeight
    InitClasses_IfNeeded = True
End Function

Function TermForms()
    
    Set frmEvents = Nothing
    
End Function

Function InitForms()
    Load frmVistaMenu
    Load frmEvents

    frmEvents.Test
End Function

Function ShowLoadingForm()

    Set g_winLoading = New frmAbout
    
    With g_winLoading
        .Caption = ""
    
        .Label3.fontSize = 9
        .Label2.fontSize = 9
        .Label1.fontSize = 9
    
        .cmdUpdate.Visible = False
        .CmdOK.Visible = False
        '.lblLink.Visible = False
        .lblLink.Top = 100 * Screen.TwipsPerPixelY
        .lblLink.Left = 90 * Screen.TwipsPerPixelX
        
        .Label5.Visible = False
        .Label6.Visible = False
        .Label4.Visible = False
        
        .lblBottom.Caption = ""
        .lblBottom.Top = 85 * Screen.TwipsPerPixelY
        .lblBottom.Width = 263 * Screen.TwipsPerPixelX
        
        .lblEmail.Caption = ""
        .lblEmail.Top = 133 * Screen.TwipsPerPixelY
        .lblEmail.FontBold = True
        .timUpdate.Enabled = True
        
        .Height = 125 * Screen.TwipsPerPixelY
        .Width = 266 * Screen.TwipsPerPixelX
        
        .Show
    End With

End Function

Function SetLoadingCaption(strCaption)

    With g_winLoading
        .lblBottom.Caption = strCaption
    End With

End Function

Sub MakeUserRollover(szResourcesPath As String)

Dim BITMAP As New GDIPBitmap
Dim BitmapGraphics As New GDIPGraphics

Dim UserPic As New GDIPImage
Dim UserFrame As New GDIPImage
Dim encoder As New GDIPImageEncoderList
Dim sBmpUserPath As String
Dim ProgramDataPath As String
Dim thisCLSID As clsid
Dim theImageFace As New GDIPImage

    ProgramDataPath = Registry.Read("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common AppData")
    sBmpUserPath = ProgramDataPath & "\Microsoft\User Account Pictures\" & Environ$("USERNAME") & ".bmp"
    If FileExists(sBmpUserPath) = False Then
        'sBmpUserPath = ProgramDataPath & "\Microsoft\User Account Pictures\user.bmp"
        sBmpUserPath = Environ$("TEMP") & "\" & Environ$("USERNAME") & ".bmp"
        
        If FileExists(GetUserTilePath) Then
            sBmpUserPath = GetUserTilePath()
        End If
        
        If FileExists(sBmpUserPath) = False Then
            If FileExists(Environ$("programdata") & "\Microsoft\User Account Pictures\user.png") Then
                sBmpUserPath = Environ$("programdata") & "\Microsoft\User Account Pictures\user.png"
            ElseIf FileExists(Environ$("programdata") & "\Microsoft\User Account Pictures\guest.bmp") Then
                sBmpUserPath = Environ$("programdata") & "\Microsoft\User Account Pictures\guest.bmp"
            ElseIf FileExists(ProgramDataPath & "\Microsoft\User Account Pictures\Guest.bmp") Then
                sBmpUserPath = ProgramDataPath & "\Microsoft\User Account Pictures\Guest.bmp"
            End If
        End If
    End If

    If FileExists(sBmpUserPath) = False Or _
        FileExists(szResourcesPath & "userframe.png") = False Then
        
        Exit Sub
    End If

    UserPic.FromFile sBmpUserPath
    
    UserFrame.FromFile szResourcesPath & "userframe.png"
    BITMAP.CreateFromSizeFormat UserFrame.Width, UserFrame.Height, GDIPlusWrapper.Format32bppArgb
    
    
    BitmapGraphics.FromImage BITMAP.Image

    BitmapGraphics.DrawImage UserPic, 11, 11, 48, 48
    BitmapGraphics.DrawImage UserFrame, 0, 0, UserFrame.Width, UserFrame.Height
    
    Set MainHelper.g_rolloverImage = BITMAP.Image.Clone

End Sub

Sub InitRecentPrograms()

Dim defaultFont As ViFont
Dim m_recentPrograms As frmFreq

    Set defaultFont = New ViFont
    Set m_recentPrograms = New frmFreq
    
    defaultFont.Size = 9
    defaultFont.Colour = vbBlack
    defaultFont.Face = g_DefaultFont.FontFace
    
    m_recentPrograms.TrueFont = defaultFont
    
    m_recentPrograms.BackColor = CLng(Layout.FrequentProgramsMenuColour)

    m_recentPrograms.Show
End Sub

Private Function HandleWindows8Utility() As Long
    
    HandleWindows8Utility = EXIT_PROGRAM
    DetermineWindowsVersion_IfNeeded
    
    If g_Windows8 Then
        
        Set MetroUtility = New Windows8Utility
        MetroUtility.ActionSettings
    End If
    
    WaitUntilDesktopIsAvailable
    Exit Function
Handler:
    LogError Err.Description
End Function

Function DetermineProgramAction() As Long

Dim newThemeName As String
Dim newOrbName As String
Dim cmdLineArguements() As String

    If CmdLine Is Nothing Then
        Set CmdLine = New CommandLine
    End If
    
    If CmdLine.Arguments > 0 Then
        Select Case LCase$(CmdLine.Argument(1))
        
            Case "/ignorepreviousinstance"
                m_IgnorePreviousInstance = True
        
            Case "/debug":
                sVar_bDebugMode = True
                
            Case "/nuke_metro"
                DetermineProgramAction = HandleWindows8Utility()
                
            Case "/install_orb"
                SetVars_IfNeeded
                SetDefaultFont_IfNeeded
            
                If CmdLine.Arguments > 1 Then
                
                    If Not InstallOrb(CmdLine.Argument(2), newOrbName) Then
                        ExitApplication
                        Exit Function
                    End If
                        
                    If FindWindow(vbNullString, MASTERID) <> 0 Then
                        DetermineProgramAction = EXIT_PROGRAM
                    
                        SendAppMessage 0, FindWindow(vbNullString, MASTERID), "NEW_ORB " & URLEncode(newOrbName)
                        ExitApplication
                        Exit Function
                    Else
                        Set Settings = New ViSettings
                        Settings.CurrentOrb = newOrbName
                    End If
                End If
            
            Case "/install_theme"
                SetVars_IfNeeded
                SetDefaultFont_IfNeeded
            
                If CmdLine.Arguments > 1 Then
                    If Not InstallTheme(CmdLine.Argument(2), newThemeName) Then
                        ExitApplication
                        Exit Function
                    End If
                        
                    If FindWindow(vbNullString, MASTERID) <> 0 Then
                        DetermineProgramAction = EXIT_PROGRAM
                    
                        SendAppMessage 0, FindWindow(vbNullString, MASTERID), "NEW_THEME " & URLEncode(newThemeName)
                        ExitApplication
                        Exit Function
                    Else
                    
                        Set Settings = New ViSettings
                        Settings.CurrentSkin = newThemeName
                    End If
                End If
                
            Case "/pin"
                SetVars_IfNeeded
                SetDefaultFont_IfNeeded
            
                If CmdLine.Arguments > 1 Then
                        
                    If FindWindow(vbNullString, MASTERID) <> 0 Then
                        DetermineProgramAction = EXIT_PROGRAM
                    
                        SendAppMessage 0, FindWindow(vbNullString, MASTERID), "PIN_FILE " & URLEncode(CmdLine.Argument(2))
                        ExitApplication
                        Exit Function
                    End If
                End If

        End Select
    End If
    
End Function

Sub Main()

    If Not InitClasses_IfNeeded Then
        Exit Sub
    End If

    If DetermineProgramAction = EXIT_PROGRAM Then
        ExitApplication
        Exit Sub
    End If
    
    If Not m_IgnorePreviousInstance And _
        FindWindow(vbNullString, MASTERID) <> 0 Then
        
        ExitApplication
        Exit Sub
    End If

    SetVars_IfNeeded
    SetDefaultFont_IfNeeded
    
    DetermineWindowsVersion_IfNeeded
    InitializeGDIIfNotInitialized
    
    g_WDSInitialized = WDSAvailable

    If Settings Is Nothing Then Set Settings = New ViSettings
    OptionsHelper.GetOptions

    If Settings.CurrentSkin <> vbNullString Then
        g_resourcesPath = sCon_AppDataPath & "_skins\" & Settings.CurrentSkin & "\"
    End If
    
    If Settings.CurrentRollover <> vbNullString Then
        g_rolloverPath = sCon_AppDataPath & "_rollover\" & Settings.CurrentRollover & "\"
        Else
                g_rolloverPath = sCon_AppDataPath & "_skins\" & Settings.CurrentSkin & "\rollover\"
    End If
        
    If ValidateOptions = False Then
        End
    End If

    If Not FileCheck(ResourcesPath) Then
        MessageBox 0, "Unable to locate resources!", "File check failed!", MB_ICONEXCLAMATION
        AppLauncherHelper.ShellEx "http://www.lee-soft.com/vistart/"
        
        Exit Sub
    End If
    
    'testComponent
    'Exit Sub
    
    
    WaitForDesktop

    App.TaskVisible = False

    'We need this
    Set g_colSearch = New Collection
    
    ProgramIndexingHelper.Initialize

    MakeUserRollover ResourcesPath
    NTEnableShutDown ""
    
    If App.CompanyName <> "Lee-Soft.com" Then
        End
    End If

    g_sVar_Layout_BackColour = GetProgramMenuBackColour()

    'If Not g_Windows8 Then
    While IsWindow(ShellHelper.g_lnghwndTaskBar) = APIFALSE And g_startButtonFindAttempts < 10
        g_startButtonFindAttempts = g_startButtonFindAttempts + 1
        
        Sleep 1000
        ShellHelper.UpdateHwnds
        DoEvents
    Wend
    'Else
    If IsWindow(ShellHelper.g_hwndStartButton) = APIFALSE Then
        If IsWindow(ShellHelper.g_lngHwndViOrbToolbar) = APIFALSE Then
            frmInstall.Show vbModal
        End If
    End If

    
        If Settings.ShowSplashScreen = True Then
                ShowLoadingForm
                SetLoadingCaption "LOADING"
        End If
         

    DoEvents
    
    InitForms
        
        If Settings.ShowSplashScreen = True Then
                g_winLoading.timSplashMin.Enabled = True
    End If
        
    If Not g_WDSInitialized Then Index_MyDirectory
    
    If App.CompanyName <> "Lee-Soft.com" Then
        End
    End If
    
    CreateFileAssociation ".vistart-theme", "VistartTheme", "A ViStart Theme Package", App.Path & "\" & App.EXEName & ".exe"
    'Exit Sub

    RegisterAppRestart
End Sub

Function GetSystemLargeIconSize() As Integer

Dim strTemp As String

    On Error GoTo Handler

    'Get this system's Icon size
    strTemp = Registry.Read("HKCU\Control Panel\Desktop\WindowMetrics\Shell Icon Size", "32")
    GetSystemLargeIconSize = CInt(strTemp)
    
    Exit Function
Handler:
    GetSystemLargeIconSize = 32

End Function

Sub ExitApplication()
    If g_Exiting Then Exit Sub
    On Error Resume Next
    
    g_Exiting = True
    
    PutOptions
    Unload frmEvents
    
    While Forms.count > 0
        Dim F As Form
        For Each F In Forms
            Unload F
        Next
    Wend
    
    DoEvents
End Sub
