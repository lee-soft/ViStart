Attribute VB_Name = "StartBarSupport"
Public Const MAX_PATH As Long = 260
   
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Enum PowerMenuCommands
    ShowOptions = 1
    ShowAbout = 2
    LogOff = 3
    PowerOff = 4
    Reboot = 5
    Hibernate = 6
    StandBy = 7
End Enum

Private Type StartOption
    Caption As String
    Shell As String
    Exists As Boolean
    ContextMenu As ContextMenu
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type rolloverImage
    Rollover As String
    Exists As Boolean
End Type

Public g_bStartMenuVisible As Boolean

Public win_txtSearch As Long

Public StartOptions() As StartOption

Public sVar_bDebugMode As Boolean
Public sVar_sFontName As String

Public sVar_Reg_StartMenu_MyDocuments As String
Public sVar_Reg_StartMenu_CommonUser As String
Public sVar_Reg_StartMenu_CurrentUser As String

Public sVar_Reg_StartMenu_CommonPrograms As String
Public sVar_Reg_StartMenu_CurrentUserPrograms As String
Public sVar_Reg_StartMenu_CurrentUserRecentItems As String
Public sVar_Reg_Desktop As String

Public g_sVar_Layout_BackColour As Long

Public iLastFileCount(0 To 3) As Long
Public iLastFolderCount(0 To 3) As Long
Public fFolder_Monitor(0 To 3) As Scripting.Folder

Public rRealStartBarPosition As RECT

'Public ProgramDB As New clsProgramDB

Public Const sCon_Reg_AppPath As String = "Software\ViStart\"

Public sCon_AppDataPath As String
Public sCon_OrbFolderPath As String

Public FSO As New FileSystemObject
Private m_optionsPopulated As Boolean

Public g_DefaultFont As GDIFont
Public g_DefaultFontItalic As GDIFont

Public g_KeyboardMenuState As Long
Public g_KeyboardSide As Long

Public AppPath As String

Dim bHasRecentItems As Boolean

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        Set m_logger = LogManager.GetLogger("StartBarSupport")
    End If
    
    Set Logger = m_logger
End Property

Public Function PopulateUserStringsFromXML(ByVal szSourceFile As String)

Dim xmlLanguageFile As New DOMDocument
    
    'Defaults
    UserVariable.Add "Log Off", "strLogOff"
    UserVariable.Add "Switch User", "strSwitchUser"
    UserVariable.Add "Shutdown", "strShutdown"
    UserVariable.Add "Restart", "strRestart"
    UserVariable.Add "Stand By", "strStandBy"
    UserVariable.Add "Hibernate", "strHibernate"
    UserVariable.Add "Exit", "strExit"
    UserVariable.Add "About", "strAbout"
    UserVariable.Add "Options", "strOptions"
    
    UserVariable.Add "All Programs", "strAllPrograms"
    UserVariable.Add "Back", "strBack"
    
    UserVariable.Add "No programs match the search criteria", "strNotFound"
    UserVariable.Add "See all results", "strSeeAllResults"
    UserVariable.Add "Start Search", "strStartSearch"
    
    UserVariable.Add "Program Options", "strProgramOptions"

    UserVariable.Add "Enable Auto-Click feature", "strEnableAutoClick"
    UserVariable.Add "Show ViStart's Tray Icon", "strShowTrayIcon"
    UserVariable.Add "Start with Windows", "strStartWithWindows"
    UserVariable.Add "Show splash screen on startup", "strSplash"
    UserVariable.Add "Clear frequently used program list", "strClearFrequentList"
    UserVariable.Add "Indexing Options", "strIndexOptions"
    UserVariable.Add "Invoke ViStart with Windows Key", "strFilterWinKey"
    
    UserVariable.Add "Pick a new Start Menu Skin...", "strNewSkin"
    UserVariable.Add "Pick a new Start Button image...", "strOrbNew"
    UserVariable.Add "Reset Orb image", "strOrbReset"
    UserVariable.Add "Run...", "strRun"
    UserVariable.Add "Sleep...", "strSleep"
    
    UserVariable.Add "OK", "strOK"
    UserVariable.Add "Cancel", "strCancel"
    UserVariable.Add "Browse", "strBrowse"
    
    UserVariable.Add "Explore", "strExplore"
    UserVariable.Add "Manage", "strManage"
    UserVariable.Add "Search", "strSearch"
    UserVariable.Add "Show on Desktop", "strShowOnDesktop"
    UserVariable.Add "Hide from Desktop", "strHideFromDesktop"
    UserVariable.Add "Show in Computer", "strShowInComputer"
    UserVariable.Add "Hide from Computer", "strHideFromComputer"
    UserVariable.Add "Open", "strOpen"
    UserVariable.Add "Don't show option in navigation pane", "strHideOption"
    UserVariable.Add "Don't pop out folder contents", "strDontPopOut"
    UserVariable.Add "Pop out folder contents", "strPopOut"
    UserVariable.Add "Rename", "strRename"
    UserVariable.Add "Send to Desktop", "strCopyToDesktop"
    UserVariable.Add "Send to ViPad", "strCopyToViPad"
    UserVariable.Add "Properties", "strProperties"
    
    UserVariable.Add "Collapse", "strCollapse"
    UserVariable.Add "Expand", "strExpand"
    UserVariable.Add "Run as administrator", "strRunAsAdmin"
    UserVariable.Add "Unpin from Start Menu", "strUnpinToStartMenu"
    UserVariable.Add "Pin To Start Menu", "strPinToStartMenu"
    UserVariable.Add "Remove from this list", "strRemoveFromList"

    UserVariable.Add "Programs", "strPrograms"
    UserVariable.Add "Files", "strFiles"
        
    UserVariable.Add "All files", "strAllExtensions"
    
    UserVariable.Add "Documents", "strDocuments"
    UserVariable.Add "Pictures", "strPictures"
    UserVariable.Add "Music", "strMusic"
    UserVariable.Add "Videos", "strVideos"
    UserVariable.Add "Games", "strGames"
    UserVariable.Add "Recent", "strRecent"
    UserVariable.Add "Computer", "strComputer"
    UserVariable.Add "Network", "strNetwork"
    UserVariable.Add "Connect To", "strConnectTo"
    UserVariable.Add "Control Panel", "strControlPanel"
    UserVariable.Add "Help and Support", "strHelp"
    UserVariable.Add "Printers and Faxes", "strPrinters"
    UserVariable.Add "Set Program Access and Defaults", "strSetDefaults"
    UserVariable.Add "Libraries", "strLibraries"
    UserVariable.Add "Downloads", "strDownloads"
    UserVariable.Add "3D Objects", "strObjects"

    UserVariable.Add "ViStart Control panel", "strViStartControlPanel"
        
    UserVariable.Add "Style", "strStyle"
    UserVariable.Add "Configure", "strConfigure"
    UserVariable.Add "Desktop", "strDesktop"

    UserVariable.Add "Start Menu Skin", "strWhichStartMenu"
    UserVariable.Add "Install...", "strInstall"
    UserVariable.Add "Select a new ViStart theme file", "strViStartTheme"
        
    UserVariable.Add "Start Orb Skin", "strWhatStarOrb"
    UserVariable.Add "Use Skin default Orb", "strSkinDefaultOrb"
    UserVariable.Add "Pick image...", "strPick"
    UserVariable.Add "Choose new Start Button image", "strViStartOrb"

    UserVariable.Add "Rollover Skin", "strWhatRollover"
    UserVariable.Add "Use Skin default Rollover", "strSkinDefaultRollover"

    UserVariable.Add "Visibility settings", "strWhatToSee"
    UserVariable.Add "Default settings for Start menu items", "strWhatToSeeOnRight"

    UserVariable.Add "Show program menu first", "strProgramsFirst"
    UserVariable.Add "Show user picture", "strShowUserPicture"

    UserVariable.Add "Don't show item", "strDontShowItem"
    UserVariable.Add "Display item as link", "strDisplayAsLink"
    UserVariable.Add "Display item as menu", "strDisplayAsMenu"

    UserVariable.Add "Set default desktop actions", "strDesktopSettings"
    UserVariable.Add "Both Windows Keys show ViStart", "strBothWinKeysViStart"
    UserVariable.Add "[Left Windows Key] shows ViStart", "strLeftWinKey"
    UserVariable.Add "[Right Windows Key] shows ViStart", "strRightWinKey"
    UserVariable.Add "Both Windows keys shows Windows Menu", "strBothWinKeys"

    UserVariable.Add "Start button shows ViStart", "strStartViStart"
    UserVariable.Add "Start button shows the Windows menu", "strStartWinMenu"

    UserVariable.Add "Restore Windows Start Menu Shortcut", "strRestoreStartMenu"
        
    UserVariable.Add "Windows 8 exclusive features defaults", "strW8Features"
        
    UserVariable.Add "Disable all Windows 8 hot corners", "strHotCorners"
    UserVariable.Add "Disable CharmsBar", "strDisableCharmsBar"
    UserVariable.Add "Disable Drag to close", "strDisableDragToClose"
    UserVariable.Add "Disable bottom left (Start) hot corner", "strDisableBottomLeftCorner"
    UserVariable.Add "Automatically go to desktop when I log in", "strSkipMetroScreen"
    UserVariable.Add "Windows 8 related features require a restart to take effect", "strW8FeaturesWarning"
        
    UserVariable.Add "(ViStart the program itself is created by Lee Matthew Chantrey)", "strCopyright"

    
    If g_Windows8 Or g_Windows81 Then
        UserVariable.Add "Metro", "startmenu"
    Else
        UserVariable.Add "Start Menu", "startmenu"
    End If

    UserVariable.Add App.Path, "apppath"
    
    UserVariable.Add ShellHelper.GetFolderPathVB(5), "CSIDL_PERSONAL"
    UserVariable.Add ShellHelper.GetFolderPathVB(39), "CSIDL_MYPICTURES"
    UserVariable.Add ShellHelper.GetFolderPathVB(&HD), "CSIDL_MYMUSIC"
    UserVariable.Add ShellHelper.GetFolderPathVB(&HE), "CSIDL_MYVIDEO"
    UserVariable.Add ShellHelper.GetFolderPathVB(&H8), "CSIDL_RECENT"
    UserVariable.Add ShellHelper.GetFolderPathVB(18), "CSIDL_NETWORK"
    
    If Not xmlLanguageFile.Load(szSourceFile) Then
        Exit Function
    End If
    
    XML_PopulateStrings xmlLanguageFile.firstChild
End Function

Public Function ShowNormalWindowsMenu()

    SetForegroundWindow g_lnghwndTaskBar
        
    g_ignoreHook = True
    SetKeyDown VK_LWINKEY
    SetKeyUp VK_LWINKEY
    g_ignoreHook = False

End Function

Public Function IsRectDifferent(ByRef rect1 As RECT, ByRef rect2 As RECT) As Boolean

    If rect1.Left <> rect2.Left Or _
        rect1.Right <> rect2.Right Or _
        rect1.Top <> rect2.Top Or _
        rect1.Bottom <> rect2.Bottom Then
        
        IsRectDifferent = True
    End If

End Function

Public Function MakeLayerdWindow(ByRef sourceForm As Form) As LayerdWindowHandles

Dim srcPoint As POINTL
Dim winSize As SIZEL

Dim mDC As Long
Dim tempBI As BITMAPINFO
Dim curWinLong As Long

Dim mainBitmap As Long
Dim oldBitmap As Long

Dim theHandles As New LayerdWindowHandles

   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32    ' Each pixel is 32 bit's wide
      .biHeight = sourceForm.ScaleHeight  ' Height of the form
      .biWidth = sourceForm.ScaleWidth    ' Width of the form
      .biPlanes = 1   ' Always set to 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
   End With

    mDC = CreateCompatibleDC(sourceForm.hdc)
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)

    If mainBitmap = 0 Then
        MsgBox "CreateDIBSection Failed", vbCritical
    End If
    oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected

    If oldBitmap = 0 Then
        MsgBox "SelectObject Failed", vbCritical
    End If

   curWinLong = GetWindowLong(sourceForm.hWnd, GWL_EXSTYLE)
   
    If SetWindowLong(sourceForm.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED Or WS_EX_TOOLWINDOW) = 0 Then
        'Logger.Error "Failed to create layered window", "Startbar_Support"
        'Exit Function
    End If

   ' Needed for updateLayeredWindow call
   srcPoint.X = 0
   srcPoint.Y = 0
   winSize.cx = sourceForm.ScaleWidth
   winSize.cy = sourceForm.ScaleHeight
    
   theHandles.mainBitmap = mainBitmap
   theHandles.oldBitmap = oldBitmap
   theHandles.theDC = mDC
   
   theHandles.SetSize winSize
   theHandles.SetPoint srcPoint
   'theHandles.
   
   Set MakeLayerdWindow = theHandles
    
End Function

Sub SendKeyToSearchBox(lngKeyCode As Long)

    If Not (lngKeyCode = vbKeyUp Or _
                lngKeyCode = vbKeyDown Or _
                lngKeyCode = vbKeyRight Or _
                lngKeyCode = vbKeyLeft) Then
    
        'frmStartMenuBase.SearchBox_Focus
        txtSearch.clicked
    End If

End Sub

Sub SetDefaultFont_IfNeeded()
    If Not g_DefaultFont Is Nothing Then Exit Sub

    Set g_DefaultFont = New GDIFont
    Set g_DefaultFontItalic = New GDIFont

    If FontExists("Tahoma") Then
        OptionsHelper.PrimaryFont = "Tahoma"
        OptionsHelper.SecondaryFont = "Tahoma"
        sVar_sFontName = "Tahoma"
        
        If FontExists("Segoe UI") Then
        
            OptionsHelper.PrimaryFont = "Segoe UI"
            sVar_sFontName = "Segoe UI"
        End If
        
    Else
        MsgBox "A compatible font was not found. Please install Tahoma and Segoe UI", vbCritical
    End If

    g_DefaultFont.Constructor OptionsHelper.PrimaryFont
    g_DefaultFontItalic.Constructor OptionsHelper.PrimaryFont, , APITRUE
End Sub

Sub SetVars_IfNeeded()
    If sCon_AppDataPath <> vbNullString Then
        Exit Sub
    End If
    
    AppPath = App.Path
    If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

    sCon_AppDataPath = Environ$("appdata") & "\ViStart\"

    If Not FSO.FolderExists(sCon_AppDataPath) Then
        If FSO.FolderExists(App.Path & "\_skins\") Then
                ' ViStart %APPDATA% folder doesn't exist _skins are present in same directory
                sCon_AppDataPath = App.Path
        End If
    End If
        
    If Not FSO.FolderExists(sCon_AppDataPath) Then
        FSO.CreateFolder sCon_AppDataPath
        
        If Err Then
            MsgBox Err.Description, vbCritical
            
            ExitApplication
            End
            Exit Sub
        End If
    End If
    
    Dim currentUserShellFoldersRegKey As RegistryKey
    Dim localMachineShellFoldersRegKey As RegistryKey
    
    Set currentUserShellFoldersRegKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
    Set localMachineShellFoldersRegKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
    
    If Not currentUserShellFoldersRegKey Is Nothing Then
    
        sVar_Reg_StartMenu_MyDocuments = currentUserShellFoldersRegKey.GetValue("Personal")
        If (LenB(sVar_Reg_StartMenu_MyDocuments) = 0) Then
            MsgBox "RegFail: My Documents Shell Folder not found", vbCritical
            End
        End If
        
        sVar_Reg_StartMenu_CurrentUser = currentUserShellFoldersRegKey.GetValue("Start Menu")
        If (LenB(sVar_Reg_StartMenu_CurrentUser) = 0) Then
            MsgBox "RegFail: Current User User Start Menu not found", vbCritical
            End
        End If
        
        sVar_Reg_StartMenu_CurrentUserPrograms = currentUserShellFoldersRegKey.GetValue("Programs")
        If (LenB(sVar_Reg_StartMenu_CurrentUserPrograms) = 0) Then
            MsgBox "RegFail: Start Menu Current User Programs not found", vbCritical
            End
        End If
        
        sVar_Reg_StartMenu_CurrentUserRecentItems = currentUserShellFoldersRegKey.GetValue("Recent")
        If (LenB(sVar_Reg_StartMenu_CurrentUserRecentItems) = 0) Then
            bHasRecentItems = False
        Else
            bHasRecentItems = True
        End If
    
    End If
    
    If Not localMachineShellFoldersRegKey Is Nothing Then
    
        sVar_Reg_StartMenu_CommonUser = localMachineShellFoldersRegKey.GetValue("Common Start Menu")
        If (LenB(sVar_Reg_StartMenu_CommonUser) = 0) Then
            MsgBox "RegFail: Common User Start Menu not found", vbCritical
            End
        End If
        
        sVar_Reg_StartMenu_CommonPrograms = localMachineShellFoldersRegKey.GetValue("Common Programs")
        If (LenB(sVar_Reg_StartMenu_CommonPrograms) = 0) Then
            MsgBox "RegFail: Start Menu Common Programs not found", vbCritical
            End
        End If
        
    End If

Dim userShellFoldersRegKey As RegistryKey
Dim userProfileDesktopKeyValue As String

    Set userShellFoldersRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")

    If Not userShellFoldersRegKey Is Nothing Then
        userProfileDesktopKeyValue = userShellFoldersRegKey.GetValue("Desktop", "%userprofile%\desktop")
    Else
        userProfileDesktopKeyValue = "%userprofile%\desktop"
    End If
    
    sVar_Reg_Desktop = VarScan(userProfileDesktopKeyValue)
    sCon_OrbFolderPath = sCon_AppDataPath & "_orbs\"
    If Not FSO.FolderExists(sCon_OrbFolderPath) Then
        FSO.CreateFolder sCon_OrbFolderPath
    End If
    
End Sub

Private Sub XML_PopulateStrings(ByRef ObjXML As IXMLDOMElement)

Dim objStrings As IXMLDOMElement
Dim objString As IXMLDOMElement
Dim thisObject As Object

    On Error GoTo Handler

    'Set objStrings = ObjXML.selectSingleNode("/strings")
    ' iterate its string children
    For Each thisObject In ObjXML.childNodes
        If TypeName(thisObject) = "IXMLDOMElement" Then
            Set objString = thisObject
            ' get all the strings
            If objString.nodeName = "string" Then
    
                If AttributeExists(objString, "id") = False Or _
                    AttributeExists(objString, "value") = False Then
                    
                    MsgBox "id or value not found in string", vbCritical
                    End
                Else
                    If UpdateColValue(UserVariable, CStr(objString.Attributes.getNamedItem("id").text), CStr(objString.Attributes.getNamedItem("value").text)) = False Then
                        Logger.Error "'" & CStr(objString.Attributes.getNamedItem("id").text) & "' is not a known string identifier", ""
                    End If
                End If
            End If
        End If
    Next
    
    Exit Sub
Handler:
    Logger.Error Err.Description, "XML_PopulateStrings"
End Sub

Public Function AttributeExists(ByRef objElem As MSXML2.IXMLDOMElement, ByVal sAttribName As String) As Boolean

Dim i As Integer
Dim objAttribs As IXMLDOMAttribute

    For Each objAttribs In objElem.Attributes

        If objAttribs.Name = sAttribName Then
            AttributeExists = True
            Exit Function
        End If
    Next

End Function

Public Function FolderExists(sSource As String) As Boolean
    FolderExists = FSO.FolderExists(sSource)
End Function

Public Function FileExists(sSource As String) As Boolean
   FileExists = FSO.FileExists(sSource)
End Function

Public Function GetTarget(strPath As String) As String

Dim c As Integer
Dim s As Integer
Dim J As Integer

    c = 0
    s = 0
    J = 0
    
    For m = 1 To Len(Path)
        GetChr0 = Right$(Path, m)
        GetChr1 = Left$(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then
            c = c + 1
        End If
    Next m
    For m = 1 To Len(Path)
        GetChr0 = Left$(Path, m)
        GetChr1 = Right$(GetChr0, 1)
        J = J + 1
        If GetChr1 = "\" Or GetChr1 = "/" Then
            J = 0
            s = s + 1
            If s = c Then
                GetTarget = Right$(GetChr0, m - J)
                Exit Function
            End If
        End If
    Next m
    
End Function

Public Function GetFileDateLastAccessed(ByVal theFilePath As String) As Date

Dim thisFile As Scripting.File

    If FSO.FileExists(theFilePath) = False Then
        Exit Function
    End If

    Set thisFile = FSO.GetFile(theFilePath)
    
    GetFileDateLastAccessed = CDate(thisFile.DateLastAccessed)

End Function

Public Sub ReverseArray(ByRef pvarArray As Variant)

Dim arrIndex As Long
Dim varSwapTo
Dim swapFrom As Long
Dim countBackward As Long

    If IsArrayInitialized(pvarArray) = False Then
        Exit Sub
    End If

For arrIndex = LBound(pvarArray) To UBound(pvarArray) / 2
    swapFrom = UBound(pvarArray) - countBackward
    
    varSwapTo = pvarArray(swapFrom)
    pvarArray(swapFrom) = pvarArray(arrIndex)
    pvarArray(arrIndex) = varSwapTo
    
    countBackward = countBackward + 1
Next


End Sub

' Omit plngLeft & plngRight; they are used internally during recursion
Public Sub QuickSort_FileAccessed(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    
    If IsArrayInitialized(pvarArray) = False Then
        Exit Sub
    End If
    
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    
    
    
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While GetFileDateLastAccessed(pvarArray(lngFirst)) < GetFileDateLastAccessed(varMid) And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While GetFileDateLastAccessed(varMid) < GetFileDateLastAccessed(pvarArray(lngLast)) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then QuickSort_FileAccessed pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then QuickSort_FileAccessed pvarArray, lngFirst, plngRight
End Sub

Function QuickSortNamesAscending( _
    ByVal col_sCollectionToAlphabetize As Collection, _
    ByVal iSortType As VbCompareMethod) As Collection
    
    On Error GoTo Handler
    
    Dim iEachItem As Long
    Dim col_sSorted As New Collection
    Dim fOrderNotChanged As Boolean
    Dim arr_sSortItems() As String
    Dim sFirstString As String
    Dim iNumberOfItems As Long
    
    iNumberOfItems = col_sCollectionToAlphabetize.count
    
    If iNumberOfItems = 0 Then
        Set QuickSortNamesAscending = col_sCollectionToAlphabetize
        Exit Function
    End If
    
    'convert to an array
    ReDim arr_sSortItems(iNumberOfItems - 1)


    For iEachItem = 1 To iNumberOfItems
        arr_sSortItems(iEachItem - 1) = col_sCollectionToAlphabetize(iEachItem)
    Next iEachItem

    

    Do


        fOrderNotChanged = True

            For iEachItem = 1 To iNumberOfItems - 1


                If Strings.StrComp(Split(arr_sSortItems(iEachItem - 1), "*")(0), Split(arr_sSortItems(iEachItem), "*")(0), iSortType) = 1 Then
                    'swap the values
                    sFirstString = arr_sSortItems(iEachItem - 1)
                    arr_sSortItems(iEachItem - 1) = arr_sSortItems(iEachItem)
                    arr_sSortItems(iEachItem) = sFirstString


                    fOrderNotChanged = False
                    End If

                Next iEachItem

                'do this until no changes were needed
            Loop Until fOrderNotChanged

            'convert back to an array

            For iEachItem = 0 To iNumberOfItems - 1
                col_sSorted.Add arr_sSortItems(iEachItem)
            Next iEachItem

            
            Set QuickSortNamesAscending = col_sSorted
            
            Exit Function
Handler:
    Logger.Error Err.Description, "QuickSortNamesAsc"
End Function

Public Function UpdateCol(ByRef Col As Collection, index, vUpdate, Optional sKey As String) As Boolean
    'Updates Collection Key, Keeps Numerical Index intact

    On Error GoTo Handler

    Col.Remove index
    
    If Col.count + 1 > index Then
        Col.Add vUpdate, sKey, index
    Else
        Col.Add vUpdate, sKey
    End If
    
    UpdateCol = True
    Exit Function
Handler:
    UpdateCol = False

End Function

Public Function IsObjectSet(AnObject As Object) As Boolean
Dim X As String
' Returns true if an object variable is initialized
' you figure this out

X = TypeName(AnObject)
If X = "Nothing" Then
  IsObjectSet = False
Else
  IsObjectSet = True
End If
End Function

Public Function ExistCol(ByRef Col As Collection, index) As Boolean
    'Updates Collection Key, Keeps Numerical Index intact

    On Error GoTo Handler

    If IsObject(Col(index)) Then
    End If

    ExistCol = True
    Exit Function
Handler:
    ExistCol = False

End Function

Public Function UpdateColValue(ByRef Col As Collection, index, vUpdate) As Boolean
    'Updates Collection Key, Keeps Numerical Index intact

    On Error GoTo Handler

    Col.Remove index
    Col.Add vUpdate, index

    UpdateColValue = True
    Exit Function
Handler:
    UpdateColValue = False

End Function

Public Function ObjectCaptionString(ByRef objSource As Object) As String

Dim strOutput As String

    On Error Resume Next

    strOutput = "[Object#2];"
    strOutput = strOutput & "Caption:" & objSource.Caption & "-"
    
    ObjectCaptionString = strOutput

End Function

Public Function ObjectChildrenCount(ByRef objSource As Object) As String

Dim strOutput As String

    On Error Resume Next

    strOutput = "[Object#3];"
    strOutput = strOutput & "Children:" & objSource.Children.count & "-"
    
    ObjectChildrenCount = strOutput

End Function

Public Function ObjectPathNameToString(ByRef objSource As Object) As String

Dim strOutput As String

    On Error Resume Next

    strOutput = "[Object#1];"
    strOutput = strOutput & "Name:" & objSource.Name & "-"
    strOutput = strOutput & "Path:" & objSource.Path
    
    ObjectPathNameToString = strOutput

End Function

Public Function GetPublicString(ByVal stringID As String, Optional Default As String = vbNullString)
    On Error GoTo Handler
    
    GetPublicString = UserVariable(stringID)
    Exit Function
    
Handler:
    GetPublicString = Default
End Function

Public Function FontExists(FontName As String) As Boolean
    Dim oFont As New StdFont
    Dim bAns As Boolean
        oFont.Name = FontName
        bAns = StrComp(FontName, oFont.Name, vbTextCompare) = 0
        FontExists = bAns
End Function

Public Function MouseInsideWindow(hWnd As Long) As Boolean

Dim WinRect As win.RECT
Dim cursorPosition As win.POINTL

    GetWindowRect hWnd, WinRect
    GetCursorPos cursorPosition
    
    If IsWindowVisible(hWnd) = APITRUE Then
        If cursorPosition.X > WinRect.Left And _
            cursorPosition.Y > WinRect.Top And _
             cursorPosition.Y < WinRect.Bottom And _
              cursorPosition.X < WinRect.Right Then
              
              MouseInsideWindow = True
        End If
    End If

End Function


Public Function ViElementToAllPrograms(srcViElement As IXMLDOMElement) As AllProgramsText
    On Error Resume Next

Dim returnRect As New AllProgramsText

    With returnRect
        .Visible = True
    
        .Style = FontStyleRegular
        .Left = srcViElement.getAttribute("x")
        .Top = srcViElement.getAttribute("y")
        .Height = srcViElement.getAttribute("height")
        .Width = srcViElement.getAttribute("width")
        
        Select Case UCase$(srcViElement.getAttribute("style"))
        
        Case "BOLD"
            .Style = FontStyleBold
        Case "ITALIC"
            .Style = FontStyleItalic
        Case "BOLD ITALIC"
            .Style = FontStyleBoldItalic
        Case "UNDERLINE"
            .Style = FontStyleUnderline
        
        End Select
    End With
        
    Set ViElementToAllPrograms = returnRect
End Function

Public Function ViElementFromXML(srcViElement As IXMLDOMElement) As GenericViElement
    On Error Resume Next

Dim returnRect As New GenericViElement

    With returnRect
        .Visible = True
    
        .Left = srcViElement.getAttribute("x")
        .Top = srcViElement.getAttribute("y")
        .Height = srcViElement.getAttribute("height")
        .Width = srcViElement.getAttribute("width")
        .Visible = srcViElement.getAttribute("visible")
        .FontID = srcViElement.getAttribute("font")
        .BackColour = CLng(HEXCOL2RGB(srcViElement.getAttribute("backcolour")))
    End With
        
    Set ViElementFromXML = returnRect
End Function

Function EnumTaskbarChildrenToFindStartButton(ByVal lHWnd As Long, ByVal lParam As Long) _
       As Long
       
       Dim RetVal As Long
       Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
       Dim WinClass As String, WinTitle As String

       RetVal = GetClassName(lHWnd, WinClassBuf, 255)
       WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
       RetVal = GetWindowText(lHWnd, WinTitleBuf, 255)
       WinTitle = StripNulls(WinTitleBuf)
 
        If LCase$(WinClass) = "start" And LCase$(WinTitle) = "start" Then
            lParam = lHWnd
            g_hwndStartButton = lHWnd
            
            EnumChildProc = False
        Else
            EnumChildProc = True
        End If
End Function


