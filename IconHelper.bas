Attribute VB_Name = "IconHelper"
Option Explicit

'To: Clean up after our selves (destroy the icon that "ExtractIcon" created)
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
'For Drawing the icon
'To: Retrieve the icon from the .EXE, .DLL or .ICO
Public Declare Function ExtractIconW Lib "shell32.dll" (ByVal hinst As Long, ByVal lpszExeFileName As Long, ByVal nIconIndex As Long) As Long
Public Declare Function ExtractIconExW Lib "shell32.dll" _
          (ByVal lpszFile As Long, _
           ByVal nIconIndex As Long, _
           ByRef phiconLarge As Any, _
           ByRef phiconSmall As Any, _
           ByVal nIcons As Long) As Long

Private Declare Function LoadImageAsLong Lib "user32" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hinst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long
   
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long
           
'To: Draw the icon into our picture box
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'Draws ANY Path
Public Declare Function SHGetFileInfoW Lib "shell32.dll" (ByVal pszPath As Long, ByVal dwAttributes As Long, psfi As SHFILEINFOW, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long

Public Enum ICON_SIZE
    LARGE_ICON = 1
    SMALL_ICON = 2
End Enum

'Unicode Format?
Public Type SHFILEINFOW
   hIcon                As Long
   iIcon                As Long
   dwAttributes         As Long
   szDisplayName(0 To 519) As Byte
   szTypeName(0 To 159)    As Byte
End Type

Private Type SHFILEINFOA
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80 ' : type name
End Type


Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private FileInfo As SHFILEINFOW
'Private FileInfo As SHFILEINFOA
Private Const Flags As Long = SHGFI_TYPENAME Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Dim lHndSysImageList As Long

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        Set m_logger = LogManager.GetLogger("IconHelper")
    End If
    
    Set Logger = m_logger
End Property

Public Function ExtractIconEx(FileName As String, hdcDestination As Long, PixelsXY As Integer, Optional lngX As Long = 0, Optional lngY As Long = 0) As Long
Dim SmallIcon As Long

    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfoW(StrPtr(FileName), 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
        'SmallIcon = SHGetFileInfoA(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfoW(StrPtr(FileName), 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
        'SmallIcon = SHGetFileInfoA(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    End If
    
    SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, hdcDestination, lngX, lngY, ILD_TRANSPARENT)
    DestroyIcon FileInfo.hIcon
End Function

Function GetIcon(srcHdc As Long, sPath As String, lngIconIndex As Long, Optional lngX As Long = 0, Optional lngY As Long = 0)

Dim lngError As Long
Dim lngIcon As Long

    If LCase$(Right$(sPath, 4)) = ".lnk" Then
        sPath = ResolveLink(sPath)
    End If

    lngIcon = ExtractIconW(App.hInstance, StrPtr(sPath), lngIconIndex)
    lngError = DrawIconEx(srcHdc, lngX, lngY, lngIcon, 32, 32, 0, 0, 3)

    If lngIcon <> 0 Then DestroyIcon lngIcon
End Function

Public Function SHExtractIcon(ByVal szPath As String, Optional ByVal iconSize As ICON_SIZE = SMALL_ICON) As Long

Dim uFlags As Long
Dim pidl As SHFILEINFOW
Dim m_hr As Long

    uFlags = SHGFI_ICON

    If iconSize = LARGE_ICON Then
        uFlags = uFlags Or SHGFI_LARGEICON
    ElseIf iconSize = SMALL_ICON Then
        uFlags = uFlags Or SHGFI_SMALLICON
    End If

    m_hr = SHGetFileInfoW(StrPtr(szPath), 0&, pidl, Len(pidl), uFlags)
    If pidl.hIcon = 0 Then
        Logger.Error "error retrieving icon", "SHExtractIcon", szPath, iconSize
        Exit Function
    End If

    SHExtractIcon = pidl.hIcon
End Function

Function CreateSmallAlphaIcon(thePath As String) As AlphaIcon

Dim SmallIcon As Long
Dim newAlphaIcon As AlphaIcon

    SmallIcon = SHGetFileInfoW(StrPtr(thePath), 0&, FileInfo, Len(FileInfo), SHGFI_SMALLICON Or SHGFI_ICON)
    
    Set newAlphaIcon = New AlphaIcon
    newAlphaIcon.CreateFromHICON FileInfo.hIcon
    
    DestroyIcon FileInfo.hIcon
    Set CreateSmallAlphaIcon = newAlphaIcon

End Function

Function GetIconDimensions() As Long

Dim largeIcon As Long
Dim iconX As Long
Dim iconY As Long

    largeIcon = SHGetFileInfoW(StrPtr(Environ$("windir") & "\notepad.exe"), 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    ImageList_GetIconSize largeIcon, iconX, iconY
    
    'SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hdc, 0, 0, ILD_TRANSPARENT)
    DestroyIcon FileInfo.hIcon

    GetIconDimensions = iconX
End Function

Public Sub SetIcon( _
      ByVal hWnd As Long, _
      ByVal sIconResName As Variant, _
      Optional ByVal bSetAsAppIcon As Boolean = True _
   )
Dim lhWndTop As Long
Dim lHWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lHWnd = hWnd
      lhWndTop = lHWnd
      Do While Not (lHWnd = 0)
         lHWnd = GetWindow(lHWnd, GW_OWNER)
         If Not (lHWnd = 0) Then
            lhWndTop = lHWnd
         End If
      Loop
   End If
   
   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   
   If VarType(sIconResName) = vbString Then
        hIconLarge = LoadImageAsString( _
              App.hInstance, CStr(sIconResName), _
              IMAGE_ICON, _
              cx, cy, _
              LR_SHARED)
   Else
        hIconLarge = LoadImageAsLong( _
              App.hInstance, CLng(sIconResName), _
              IMAGE_ICON, _
              cx, cy, _
              LR_SHARED)
   End If
   
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cx = GetSystemMetrics(SM_CXSMICON)
   cy = GetSystemMetrics(SM_CYSMICON)
   
   If VarType(sIconResName) = vbString Then
        hIconSmall = LoadImageAsString( _
              App.hInstance, CStr(sIconResName), _
              IMAGE_ICON, _
              cx, cy, _
              LR_SHARED)
   Else
        hIconSmall = LoadImageAsLong( _
              App.hInstance, CLng(sIconResName), _
              IMAGE_ICON, _
              cx, cy, _
              LR_SHARED)
   End If
         
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub


