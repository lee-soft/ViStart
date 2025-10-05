@ -1,691 +1,691 @@
Attribute VB_Name = "ShellHelper"
Option Explicit

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long
Private Declare Function GetNextWindow Lib "user32.dll" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, ByRef pData As appBarData) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
                        
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExW" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Integer) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32" (ByVal pBSTR As Long, ByVal psz As Long, ByVal length As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (Src As Any, Dst As Any) As Long

'Declare Function GetThreadLocale Lib "kernel32" Alias "GetThreadLocale" () As Long

Private m_wScriptShellObject As Object

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type


Public g_lnghwndTaskBar As Long
Public g_lnghwndStartMenu As Long
Public g_hwndStartButton As Long
Public g_hwndReBarWindow32 As Long
Public g_ViGlanceContainer As Long
Public g_ViGlanceOrb As Long

Public g_hwndMSTask As Long
Public g_lngHwndViOrbToolbar As Long

Public g_ViGlanceOpen As Boolean

Public g_WinVer() As String
Public g_WinDecimalSeparator As String

Public g_WindowsVersion As Double
Public g_WindowsVersionFull As String

Public g_WindowsLanguageLCID As Long
Public g_WindowsLanguageCulture As String
Public g_WindowsLanguageCountry As String
Public g_WindowsLanguage As String
Public g_WindowsLanguageInt As String

Public g_WindowsXP As Boolean
Public g_WindowsVista As Boolean
Public g_Windows7 As Boolean
Public g_Windows8 As Boolean
Public g_Windows81 As Boolean
Public g_Windows10 As Boolean
Public g_Windows11 As Boolean
Public g_Windows12 As Boolean

Public g_CLSID_MYDOCS As String
Public g_CLSID_MYPIC As String
Public g_CLSID_MYMUS As String
Public g_CLSID_MYVID As String
Public g_CLSID_DOWNLOADS As String
Public g_CLSID_3DOBJECTS As String
Public g_CLSID_DESKTOP As String
Public g_CLSID_NETWORK As String
Public g_CLSID_USERPROFILE As String
Public g_CLSID_LIBRARIES As String
Public g_CLSID_COMPUTER As String
Public g_CLSID_GAMES As String
Public g_CLSID_CONTROLPANEL As String
        
Public Type appBarData
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long 'message specific
End Type

Public Enum AbeBarEnum
   abe_bottom = 3
   ABE_LEFT = 0
   ABE_RIGHT = 2
   ABE_TOP = 1
End Enum

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Const ABM_GETTASKBARPOS As Long = &H5
Private Const SHGFP_Type_CURRENT = &H0

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        Set m_logger = LogManager.GetLogger("ShellHelper")
    End If
    
    Set Logger = m_logger
End Property

Private Function PtrAdd(ByVal Address As Long, ByVal Offset As Long) As Long
' unsigned pointer arithmetic, moves overflow by toggling the sign bit
' required when using /LARGEADDRESSAWARE on 64bit windows
    Const SIGN_BIT As Long = &H80000000
    PtrAdd = (Address Xor SIGN_BIT) + Offset Xor SIGN_BIT
End Function

Private Property Get DUInt(ByVal Address As Long) As Long
' Compensate for VB's lack of unsigned types
' Copies a 16bit Unsigned Integer from a pointer into a Long
    GetMem2 ByVal Address, DUInt
End Property
' https://blogs.msdn.microsoft.com/oldnewthing/20040130-00/?p=40813/
Function FindStringResourceEx(ByVal hInstance As Long, ByVal uId As Long, ByVal langId As Long) As String

    Const STRINGS_PER_BUCKET As Long = 16&
    Const RT_STRING As Long = 6&
    Const WCHARSIZE As Long = 2

    Dim hResource As Long
    Dim hGlobal As Long
    Dim ptr As Long, i As Long

    hResource = FindResourceEx(hInstance, RT_STRING, uId \ STRINGS_PER_BUCKET + 1, langId)
        
    If hResource Then
        hGlobal = LoadResource(hInstance, hResource)
        If hGlobal Then
            ptr = LockResource(hGlobal)
            If ptr Then
                For i = 1 To uId And (STRINGS_PER_BUCKET - 1)
                    ptr = PtrAdd(ptr, (1 + DUInt(ptr)) * WCHARSIZE)
                Next
                SysReAllocStringLen VarPtr(FindStringResourceEx), PtrAdd(ptr, 2), DUInt(ptr)
            End If
        End If
    End If
End Function


Public Function GetStringFromFile(ByVal RESOURCE_FILE As String, ByVal RESOURCE_LANG As Long, ByVal STRING_ID As Long)

    'Const RESOURCE_FILE = "c:\windows\system32\shell32.dll"
    'Const RESOURCE_LANG = 1029
    'Const STRING_ID = 4148

    Dim hModule As Long
    hModule = LoadLibrary(StrPtr(RESOURCE_FILE))
        
        Dim ExtractedString As String
        
    If hModule Then
        
                ExtractedString = FindStringResourceEx(hModule, STRING_ID, RESOURCE_LANG)
                GetStringFromFile = ExtractedString
                
        FreeLibrary hModule
    End If
        
End Function


'Purpose     :  Allows the user to select a file name from a local or network directory.
'Inputs      :  sInitDir            The initial directory of the file dialog.
'               sFileFilters        A file filter string, with the following format:
'                                   eg. "Excel Files;*.xls|Text Files;*.txt|Word Files;*.doc"
'               [sTitle]            The dialog title
'               [lParentHwnd]       The handle to the parent dialog that is calling this function.
'Outputs     :  Returns the selected path and file name or a zero length string if the user pressed cancel
Function BrowseForFile(sInitDir As String, Optional ByVal sFileFilters As String, Optional sTitle As String = "Open File", Optional lParentHwnd As Long) As String
    Dim tFileBrowse As OPENFILENAME
    Const clMaxLen As Long = 5000
    
    Dim theBuffer As String
    theBuffer = String$(clMaxLen, Chr$(0))
    
    tFileBrowse.lStructSize = Len(tFileBrowse)
    
    'Replace friendly deliminators with nulls
    sFileFilters = Replace(sFileFilters, "|", vbNullChar)
    sFileFilters = Replace(sFileFilters, ";", vbNullChar)
    If Right$(sFileFilters, 1) <> vbNullChar Then
        'Add final delimiter
        sFileFilters = sFileFilters & vbNullChar
    End If
    
    'Select a filter
    tFileBrowse.lpstrFilter = StrPtr(sFileFilters & GetPublicString("strAllExtensions") & " (*.*)" & vbNullChar & "*.*" & vbNullChar)
    'create a buffer for the file
    tFileBrowse.lpstrFile = StrPtr(theBuffer)
    'set the maximum length of a returned file
    tFileBrowse.nMaxFile = clMaxLen + 1
    'Create a buffer for the file title
    tFileBrowse.lpstrFileTitle = StrPtr(Space$(clMaxLen))
    'Set the maximum length of a returned file title
    tFileBrowse.nMaxFileTitle = clMaxLen + 1
    'Set the initial directory
    tFileBrowse.lpstrInitialDir = StrPtr(sInitDir)
    'Set the parent handle
    tFileBrowse.hwndOwner = lParentHwnd
    'Set the title
    tFileBrowse.lpstrTitle = StrPtr(sTitle)
    
    'No flags
    tFileBrowse.Flags = 0

    'Show the dialog
    If GetOpenFileName(tFileBrowse) Then
        BrowseForFile = Trim$(GetString((tFileBrowse.lpstrFile)))
        If Right$(BrowseForFile, 1) = vbNullChar Then
            'Remove trailing null
            BrowseForFile = Left$(BrowseForFile, Len(BrowseForFile) - 1)
        End If
    End If
End Function

Function WindowsVersion() As Double
    DetermineWindowsVersion_IfNeeded
        
    WindowsVersion = g_WindowsVersion

    If g_Windows11 Then
        WindowsVersion = 11#
    End If
        
End Function

Function DetermineWindowsVersion_IfNeeded()
        
    If g_WindowsXP Or g_WindowsVista Or g_Windows7 Or g_Windows8 Or g_Windows81 Or g_Windows10 Or g_Windows11 Or g_Windows12 Then
        Exit Function
    End If

        g_WindowsLanguageLCID = Trim(GetUserDefaultLCID)
        g_WindowsLanguageCulture = GetLocaleCulture
        g_WindowsLanguageCountry = GetCountryName
        g_WindowsLanguage = RTrim(Replace(GetLocaleLanguage, "(" & g_WindowsLanguageCountry & ")", ""))
        g_WindowsLanguageInt = GetLocaleLanguageInt
        
        'debugprint g_WindowsLanguageLCID
        'debugprint g_WindowsLanguageCulture
        'debugprint GetCountryName
        'debugprint g_WindowsLanguage
        'debugprint g_WindowsLanguageInt
        
        g_WinDecimalSeparator = GetDecimalSeparator

    g_WindowsXP = False
    g_WindowsVista = False
    g_Windows7 = False
    g_Windows8 = False
    g_Windows81 = False
    g_Windows10 = False
    g_Windows11 = False
    g_Windows12 = False
        
    Dim kernalPath As String: kernalPath = Environ("windir") & "\System32\kernel32.dll"                         ' ntoskrnl.exe may not be accesible in certain situations
    Dim kernalFileInfo As FileVersionInfo: Set kernalFileInfo = FileVersionInfoHelper.GetVersionInfo(kernalPath)
    
    Dim currentVersionRegKey As RegistryKey
    Set currentVersionRegKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion")
    
    If currentVersionRegKey Is Nothing Then
        Logger.Error "Unable to get registry key", "DetermineWindowsVersion_IfNeeded"
        Exit Function
    End If
    
    Dim winRegistryVersion As String: winRegistryVersion = currentVersionRegKey.GetValue("CurrentVersion")
    
    g_WindowsVersion = kernalFileInfo.ProductMajorPart & g_WinDecimalSeparator & kernalFileInfo.ProductMinorPart
    
        
        If kernalFileInfo.ProductMajorPart = 5 And Not kernalFileInfo.ProductMinorPart = 0 Then
                ' XP / 2003
        g_WindowsXP = True
        ElseIf kernalFileInfo.ProductMajorPart = 6 Then
                ' Vista, 7, 8, 8.1
                If kernalFileInfo.ProductMinorPart = 0 Then
                        g_WindowsVista = True
                ElseIf kernalFileInfo.ProductMinorPart = 1 Then
                        g_Windows7 = True
                ElseIf kernalFileInfo.ProductMinorPart = 2 Then
                        g_Windows8 = True
                ElseIf kernalFileInfo.ProductMinorPart = 3 Then
                        g_Windows81 = True
                End If
        
        ElseIf kernalFileInfo.ProductMajorPart = 10 And kernalFileInfo.ProductBuildPart >= 22000 Then
                ' Windows 11
                g_Windows11 = True
        
        ElseIf kernalFileInfo.ProductMajorPart = 10 Then
                ' Windows 10
                g_Windows10 = True
        
        ElseIf kernalFileInfo.ProductMajorPart = 12 Then
                g_Windows12 = True
                
        Else
        MsgBox "This version of Windows (" & g_WindowsVersion & ") is unknown.. ViStart may not behave as expected!..", vbCritical
                
    End If
    
    g_CLSID_3DOBJECTS = "{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}"
    g_CLSID_DESKTOP = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
    
    g_CLSID_USERPROFILE = "{59031A47-3F72-44A7-89C5-5595FE6B30EE}"

    g_CLSID_LIBRARIES = "{031E4825-7B94-4dc3-B131-E946B44C8DD5}"
            
    g_CLSID_COMPUTER = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    g_CLSID_GAMES = "{ED228FDF-9EA8-4870-83b1-96b02CFE0D52}"
    
    g_CLSID_CONTROLPANEL = "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"
    
    If g_WindowsVersion < 6 Then
        g_CLSID_CONTROLPANEL = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    End If
    
    If g_WindowsVersion >= 10 Then
        ' Windows 10 +
        g_CLSID_MYDOCS = "{D3162B92-9365-467A-956B-92703ACA08AF}"
        g_CLSID_MYPIC = "{24AD3AD4-A569-4530-98E1-AB02F9417AA8}"
        g_CLSID_MYMUS = "{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE}"
        g_CLSID_MYVID = "{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A}"
        g_CLSID_DOWNLOADS = "{088E3905-0323-4B02-9826-5D99428E115F}"
        
        g_CLSID_NETWORK = "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"
    Else
        ' XP - 8.1
        g_CLSID_MYDOCS = "{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}"
        g_CLSID_MYPIC = "{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}"
        g_CLSID_MYMUS = "{1CF1260C-4DD0-4EBB-811F-33C572699FDE}"
        g_CLSID_MYVID = "{A0953C92-50DC-43BF-BE83-3742FED03C9C}"
        g_CLSID_DOWNLOADS = "{374DE290-123F-4565-9164-39C4925E467B}"
        
        g_CLSID_NETWORK = "{208D2C60-3AEA-1069-A2D7-08002B30309D}"
    End If
  
End Function

Function WaitUntilDesktopIsAvailable()

    Do
        Sleep 500
    Loop While UpdateHwnds() = False

End Function

Function UpdateHwnds() As Boolean

Dim lngHwndTaskBar As Long
Dim lngHwndStartMenu As Long

Dim bReturn As Boolean
Dim lParamReturn As Long

    bReturn = False
    lngHwndTaskBar = FindWindow("Shell_TrayWnd", "")

    If lngHwndTaskBar <> g_lnghwndTaskBar Then
        bReturn = True
        g_lnghwndTaskBar = lngHwndTaskBar
    End If
    
    g_hwndReBarWindow32 = FindWindowEx(ByVal lngHwndTaskBar, ByVal 0&, "ReBarWindow32", vbNullString)
    g_lngHwndViOrbToolbar = FindWindowEx(ByVal g_hwndReBarWindow32, ByVal 0&, "ToolbarWindow32", "Start")
    
    If g_Windows8 Or g_Windows7 Then
        g_hwndMSTask = FindWindowEx(ByVal g_hwndReBarWindow32, ByVal 0&, "MSTaskSwWClass", vbNullString)
    End If
    
    If g_WindowsXP Then
        g_hwndStartButton = FindWindowEx(lngHwndTaskBar, 0, "Button", vbNullString)
        If g_hwndStartButton = 0 Then
            'Reset update trigger (forcing routine to later update again)
            lngHwndTaskBar = -1
        End If
        
    Else
    
        g_hwndStartButton = FindWindow("Button", "Start")

        If g_hwndStartButton = 0 Then
            g_hwndStartButton = FindWindow("Button", vbNullString)
            
            If g_hwndStartButton = 0 Then
                Call EnumChildWindows(lngHwndTaskBar, AddressOf EnumTaskbarChildrenToFindStartButton, lParamReturn)
            End If
        End If

        If g_hwndStartButton = 0 Then
            'Reset update trigger (forcing routine to later update again)
            
            If g_lnghwndTaskBar > 0 And Not g_Windows8 Then lngHwndTaskBar = -1
        End If
    End If
    
    g_ViGlanceContainer = FindWindow("ThunderRT6FormDC", "Running Applications")
    If g_ViGlanceContainer <> 0 Then

        g_ViGlanceOrb = FindWindow("ThunderRT6FormDC", "#Start~ViGlance#")
        If g_ViGlanceOrb <> 0 Then
            g_ViGlanceOpen = True
        Else
            g_ViGlanceOpen = False
        End If
    Else
        g_ViGlanceOpen = False
    End If
    
    lngHwndStartMenu = FindWindow("DV2ControlHost", "Start Menu")
    If lngHwndStartMenu = 0 Then
        lngHwndStartMenu = FindWindow("DV2ControlHost", vbNullString)
    End If
    
    If lngHwndStartMenu <> g_lnghwndStartMenu Then
        'modHookReciever.SubClass 0
        
        bReturn = True
        g_lnghwndStartMenu = lngHwndStartMenu
    End If
    
    UpdateHwnds = bReturn
        
End Function

Public Function GetTaskBarEdge() As AbeBarEnum
        
Dim abd As appBarData

    abd.cbSize = LenB(abd)
    abd.hWnd = ShellHelper.g_lnghwndTaskBar
    SHAppBarMessage ABM_GETTASKBARPOS, abd
    
    GetTaskBarEdge = GetEdge(abd.rc)

End Function

Private Function GetEdge(rc As RECT) As Long

Dim uEdge As Long: uEdge = -1

    If (rc.Top = rc.Left) And (rc.Bottom > rc.Right) Then
        uEdge = ABE_LEFT
    ElseIf (rc.Top = rc.Left) And (rc.Bottom < rc.Right) Then
        uEdge = ABE_TOP
    ElseIf (rc.Top > rc.Left) Then
        uEdge = abe_bottom
    Else
        uEdge = ABE_RIGHT
    End If
    
    GetEdge = uEdge

End Function

Function IsTaskBarBehindWindow(hWnd As Long)
    
    If GetZOrder(g_lnghwndTaskBar) > GetZOrder(hWnd) Then
        IsTaskBarBehindWindow = True
    Else
        IsTaskBarBehindWindow = False
    End If
    
End Function

Function IsWindowTopMost(hWnd As Long)

Dim windowStyle As Long

    IsWindowTopMost = False
    windowStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If IsStyle(windowStyle, WS_EX_TOPMOST) Then
        IsWindowTopMost = True
    End If

End Function

Public Function IsStyle( _
      ByVal lAll As Long, _
      ByVal lBit As Long) As Boolean
      
   IsStyle = False
   If (lAll And lBit) = lBit Then
      IsStyle = True
   End If
End Function

Public Function GetZOrder(ByVal hWndTarget As Long) As Long
    
Dim hWnd As Long
Dim lngZOrder As Long

    ' Loop through window list and
    ' compare to hWnd to hwndTarget, to find global ZOrder
    hWnd = GetTopWindow(0)
    lngZOrder = 0
    
    Do While hWnd And hWnd <> hWndTarget
       ' Get next window and move on.
        hWnd = GetNextWindow(hWnd, _
          GW_HWNDNEXT)
        lngZOrder = lngZOrder + 1
    Loop
    
    GetZOrder = lngZOrder

End Function

Public Function hWndBelongToUs(hWnd As Long, Optional ExceptionHwnd As Long) As Boolean

Dim thisForm As Form
    hWndBelongToUs = False

    For Each thisForm In Forms
        If thisForm.hWnd = hWnd Then
            If hWnd = ExceptionHwnd Then
                hWndBelongToUs = False
                
                Dim t As String * 64
                GetWindowText hWnd, t, Len(t)
            Else
                hWndBelongToUs = True
            End If
            
            Exit For
        End If
    Next
    
End Function

Public Function CalculateTopBasedOnDPI(ByVal theDPI As Long, ByVal clientHeight As Long) As Long
    
Dim scalePercentage As Single
Dim taskBarHeight As Single

Dim TaskBar As win.RECT
    TaskBar = GetTaskBarPosition
    
    scalePercentage = 96 / theDPI
    taskBarHeight = (TaskBar.Bottom - TaskBar.Top) * scalePercentage
    
    CalculateTopBasedOnDPI = ((Screen.Height / Screen.TwipsPerPixelY) - taskBarHeight) - (clientHeight)
End Function

Public Function GetTaskBarPosition() As win.RECT

Dim TaskBar As appBarData

    TaskBar.hWnd = g_lnghwndTaskBar
    SHAppBarMessage ABM_GETTASKBARPOS, TaskBar
    GetTaskBarPosition = TaskBar.rc

End Function

Private Function SHGetSpecialFolderLocationVB(ByVal lFolder As Long) As String
    Dim lRet As Long, IDL As ITEMIDLIST, sPath As String

    lRet = SHGetSpecialFolderLocation(100&, lFolder, IDL)
    If lRet = 0 Then
        sPath = String$(512, Chr$(0))
        lRet = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        SHGetSpecialFolderLocationVB = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Else
        SHGetSpecialFolderLocationVB = vbNullString
    End If
End Function

Public Function GetFolderPathVB(ByVal lFolder As Long) As String
    Dim Path As String
    If IsVistaOrHigher() Then
        GetFolderPathVB = SHGetSpecialFolderLocationVB(lFolder)
    Else
        Path = Space$(MAX_PATH)
        SHGetFolderPath 0, lFolder, 0, SHGFP_Type_CURRENT, Path
        GetFolderPathVB = Left$(Path, InStr(Path, vbNullChar) - 1)
    End If
End Function

Private Function IsVistaOrHigher() As Boolean

    DetermineWindowsVersion_IfNeeded
    
    If WindowsVersion >= 6 Then
        
                IsVistaOrHigher = True
        
    End If
    
End Function

Public Function GetIconLocationAndIndex(ByVal szPath As String, ByRef iconIndex As Long, ByRef szIconLocation As String) As Boolean

Dim theLinkFile As ShellLinkObject
Dim theLinkFileA As CShellLink

    Set theLinkFile = ShellHelper.GetShellLink(szPath)
    If theLinkFile Is Nothing Then Set theLinkFileA = ShellHelper.GetShellLinkA(szPath)
    
    If theLinkFile Is Nothing And theLinkFileA Is Nothing Then
        Exit Function
    End If

    If theLinkFile Is Nothing Then
        szIconLocation = theLinkFileA.IconLocation(iconIndex)
        If szIconLocation = "" Then
            szIconLocation = theLinkFileA.Target
        End If
        
        szIconLocation = VarScan(szIconLocation)
    Else
        iconIndex = theLinkFile.GetIconLocation(szIconLocation)
        If szIconLocation = "" Then
            szIconLocation = theLinkFile.Target
        End If
        szIconLocation = VarScan(szIconLocation)
    End If

    GetIconLocationAndIndex = True
End Function

Public Function GetShellLinkA(ByVal szLinkPath As String) As CShellLink

Dim LnkFile As CShellLink
    
    Set LnkFile = New CShellLink
    If LnkFile.Load(szLinkPath) = False Then
        Exit Function
    End If

    Set GetShellLinkA = LnkFile
End Function

Public Function GetShellLink(ByVal szLinkPath As String) As ShellLinkObject
    On Error GoTo Handler
    
Dim lnk As New ShellLinkObject

    If lnk.Resolve(szLinkPath) Then
        Set GetShellLink = lnk
    End If
    
    Exit Function
Handler:
    If Err.Number <> 70 Then
        Logger.Error Err.Description & " {" & Err.Number & "}", "GetShellLink", szLinkPath
    End If
End Function

Public Function GetGlobalWScriptShellObject() As Object
    If m_wScriptShellObject Is Nothing Then
        Set m_wScriptShellObject = CreateObject("WScript.Shell")
    End If
    
    Set GetGlobalWScriptShellObject = m_wScriptShellObject
End Function


Public Function debugPrint(ByVal strInput As String)
        'Open App.Path & "\ViStart.log" For Append As #1
        Open Environ$("appdata") & "\ViStart\ViStart.log" For Append As #1
        Write #1, strInput
        Close #1
End Function