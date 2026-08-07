Attribute VB_Name = "AppsFolderHelper"
Option Explicit

'Enumerates the "All Applications" virtual shell folder (shell:AppsFolder), where
'Microsoft Store / packaged apps like Win11 Notepad and Calculator live - they have
'no entries in the classic Start Menu folder tree, so without this they are
'invisible to ViStart. Also extracts their icons, which aren't exe resources but
'manifest-driven PNG assets only reachable through IShellItemImageFactory.

Private Const APPSFOLDER_CLSID As String = "{4234d49b-0245-4df3-b780-3893943456e1}"
Private Const IID_ISHELLITEMIMAGEFACTORY As String = "{BCC18B79-BA16-442F-80C4-8A59C30C463B}"
Private Const SHELL_APPSFOLDER_PREFIX As String = "SHELL:APPSFOLDER\"

Private Const CC_STDCALL As Long = 4
Private Const SIIGBF_ICONONLY As Long = &H4

Private Type APPICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Declare Function SHGetIDListFromObject Lib "shell32" (ByVal punk As IUnknown, ppidl As Long) As Long
Private Declare Function SHCreateItemFromIDList Lib "shell32" (ByVal pidl As Long, riid As Any, ppv As Long) As Long
Private Declare Function SHParseDisplayName Lib "shell32" (ByVal pszName As Long, ByVal pbc As Long, ppidl As Long, ByVal sfgaoIn As Long, psfgaoOut As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function App_CLSIDFromString Lib "ole32" Alias "CLSIDFromString" (ByVal lpszGuid As Long, pGuid As Any) As Long
Private Declare Sub App_CoTaskMemFree Lib "ole32" Alias "CoTaskMemFree" (ByVal pv As Long)
Private Declare Function App_CreateBitmap Lib "gdi32" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByVal lpvBits As Long) As Long
Private Declare Function App_CreateIconIndirect Lib "user32" Alias "CreateIconIndirect" (piconinfo As APPICONINFO) As Long
Private Declare Function App_DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long

'PIDLs captured from the live FolderItems during enumeration, keyed by uppercase
'parsing path. Re-binding through the PIDL yields the real UWP icon where the
'parsing name alone reconstructs a thinner item with only a generic icon.
'They stay allocated for the process lifetime, like the .NET port.
Private m_pidlCache As Collection

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        Set m_logger = LogManager.GetLogger("AppsFolderHelper")
    End If

    Set Logger = m_logger
End Property

'Returns a Collection of ColItem (IsFile=True) for every app the shell knows,
'keyed by uppercase path. Win32 apps also appear in AppsFolder with their real
'file path; callers dedupe against the classic start menu walk by caption.
Public Function EnumerateApps() As Collection
    On Error GoTo Handler

Dim shellApp As Object
Dim appsFolder As Object
Dim thisItem As Object

Dim theName As String
Dim thePath As String
Dim newApp As ColItem
Dim thisPidl As Long

    Set EnumerateApps = New Collection
    If m_pidlCache Is Nothing Then Set m_pidlCache = New Collection

    Set shellApp = CreateObject("Shell.Application")
    Set appsFolder = shellApp.Namespace("shell:::" & APPSFOLDER_CLSID)
    If appsFolder Is Nothing Then Exit Function

    On Error Resume Next
    For Each thisItem In appsFolder.Items
        theName = vbNullString
        thePath = vbNullString

        theName = thisItem.Name
        thePath = thisItem.Path
        Err.Clear

        If Len(theName) > 0 And Len(thePath) > 0 Then

            'Virtual entries return a bare AppUserModelID as their path; normalise
            'to the parsing-name form that launch and icon extraction understand
            If FileExists(thePath) = False And LCase$(Left$(thePath, 6)) <> "shell:" Then
                thePath = "shell:AppsFolder\" & thePath
            End If

            If Not ExistInCol(EnumerateApps, UCase$(thePath)) Then
                Set newApp = New ColItem
                newApp.IsFile = True
                newApp.Caption = theName
                newApp.Path = thePath
                newApp.VirtualPath = thePath

                EnumerateApps.Add newApp, UCase$(thePath)

                thisPidl = 0
                If SHGetIDListFromObject(thisItem, thisPidl) = 0 And thisPidl <> 0 Then
                    If Not ExistInCol(m_pidlCache, UCase$(thePath)) Then
                        m_pidlCache.Add thisPidl, UCase$(thePath)
                    Else
                        App_CoTaskMemFree thisPidl
                    End If
                End If
            End If
        End If
    Next

    Exit Function
Handler:
    Logger.Error Err.Description, "EnumerateApps"
End Function

'HICON for a "shell:AppsFolder\<AUMID>" parsing path, or 0. Caller owns the icon.
Public Function GetAppIcon(ByVal szParsingPath As String, ByVal iconSize As Long) As Long
    On Error GoTo Handler

Dim thisPidl As Long
Dim pFactory As Long
Dim hBmp As Long
Dim iidFactory(0 To 3) As Long
Dim parseName As String
Dim dummySfgao As Long

    szParsingPath = UCase$(szParsingPath)
    If m_pidlCache Is Nothing Then Set m_pidlCache = New Collection

    If ExistInCol(m_pidlCache, szParsingPath) Then thisPidl = m_pidlCache(szParsingPath)

    'Fallback when the app hasn't been enumerated this session (e.g. a pinned
    'Store app drawn before the program list builds). SHParseDisplayName doesn't
    'accept the "shell:" convenience syntax - use the CLSID-rooted canonical form.
    If thisPidl = 0 Then
        If Left$(szParsingPath, Len(SHELL_APPSFOLDER_PREFIX)) <> SHELL_APPSFOLDER_PREFIX Then Exit Function

        parseName = "::" & APPSFOLDER_CLSID & "\" & Mid$(szParsingPath, Len(SHELL_APPSFOLDER_PREFIX) + 1)
        If SHParseDisplayName(StrPtr(parseName), 0, thisPidl, 0, dummySfgao) <> 0 Then Exit Function
        If thisPidl = 0 Then Exit Function

        m_pidlCache.Add thisPidl, szParsingPath
    End If

    App_CLSIDFromString StrPtr(IID_ISHELLITEMIMAGEFACTORY), iidFactory(0)
    If SHCreateItemFromIDList(thisPidl, iidFactory(0), pFactory) <> 0 Then Exit Function
    If pFactory = 0 Then Exit Function

    hBmp = CallGetImage(pFactory, iconSize)
    ComRelease pFactory

    If hBmp = 0 Then Exit Function

    GetAppIcon = IconFromBitmap32(hBmp, iconSize)
    App_DeleteObject hBmp

    Exit Function
Handler:
    Logger.Error Err.Description, "GetAppIcon", szParsingPath
End Function

'IShellItemImageFactory::GetImage is vtable slot 3 (offset 12). Its SIZE argument
'is passed by value, which on x86 stdcall is just two dword slots.
Private Function CallGetImage(ByVal pFactory As Long, ByVal iconSize As Long) As Long

Dim hBmp As Long
Dim hr As Long
Dim vParams(0 To 3) As Variant
Dim vTypes(0 To 3) As Integer
Dim vPtrs(0 To 3) As Long
Dim vReturn As Variant
Dim i As Long

    vParams(0) = iconSize          'SIZE.cx
    vParams(1) = iconSize          'SIZE.cy
    vParams(2) = SIIGBF_ICONONLY
    vParams(3) = VarPtr(hBmp)

    For i = 0 To 3
        vTypes(i) = vbLong
        vPtrs(i) = VarPtr(vParams(i))
    Next

    hr = DispCallFunc(pFactory, 12&, CC_STDCALL, vbLong, 4&, vTypes(0), vPtrs(0), vReturn)
    If hr = 0 Then
        If CLng(vReturn) = 0 Then CallGetImage = hBmp
    End If

End Function

'IUnknown::Release is vtable slot 2 (offset 8)
Private Sub ComRelease(ByVal pUnknown As Long)

Dim vReturn As Variant

    DispCallFunc pUnknown, 8&, CC_STDCALL, vbLong, 0&, ByVal 0&, ByVal 0&, vReturn

End Sub

'The factory hands back a 32bpp ARGB bitmap; icons need an HICON, so wrap it via
'CreateIconIndirect (which copies, so both source bitmaps are freed by callers).
Private Function IconFromBitmap32(ByVal hBmp As Long, ByVal iconSize As Long) As Long

Dim theInfo As APPICONINFO
Dim hMask As Long

    hMask = App_CreateBitmap(iconSize, iconSize, 1, 1, 0)
    If hMask = 0 Then Exit Function

    theInfo.fIcon = 1
    theInfo.hbmMask = hMask
    theInfo.hbmColor = hBmp

    IconFromBitmap32 = App_CreateIconIndirect(theInfo)
    App_DeleteObject hMask

End Function
