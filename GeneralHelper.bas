Attribute VB_Name = "GeneralHelper"
Option Explicit

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare Sub DragAcceptFiles Lib "shell32" (ByVal hWnd As Long, ByVal _
    BOOL As Long)
Public Declare Sub DragFinish Lib "shell32" (ByVal hDrop As Integer)
Public Declare Function DragQueryFileW Lib "shell32" (ByVal wParam As Long, _
    ByVal index As Long, ByVal lpszFile As Long, ByVal BufferSize As Long) _
    As Long
    
Private Declare Function InvalidateRect Lib "user32" ( _
    ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, _
ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Declare Function CreateFileW Lib "kernel32" _
  (ByVal lpFileName As Long, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long
   
Public Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TrackMouseEvent) As Long
Public Const TME_LEAVE As Long = &H2
Public Const WM_USER As Long = &H400

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Public Declare Function SendMessageW Lib "user32" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long _
) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongW" _
  (ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal wNewWord As Long) As Long

Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal _
    hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    
Public Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (ByRef oldValue As Long) As Long
Public Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByRef oldValue As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long
    
Private Declare Function RegisterApplicationRestart Lib "kernel32" (ByVal stringCmdArgs As String, Flags As Long) As Long

Private Declare Function GetUserTilePathAPI Lib "shell32.dll" Alias "#261" _
                   (ByVal theUsername As String, _
                   ByVal whatEver As Long, _
                   ByVal picPath As String, _
                   ByVal maxLength As Long) As Long
                   
Private Declare Function HashData Lib "shlwapi" (pbData As Any, ByVal cbData As Long, pbHash As Any, ByVal cbHash As Long) As Long

Public Const CONTEXT_MENU As String = "Pin to ViStart"
Public Const MASTERID As String = "ViStart_27081987_Master"

Public Const VK_LWIN As Long = &H5B

Public Const KEYEVENTF_KEYUP = &H2

Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100

Public Const ULW_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000

Public Const AC_SRC_ALPHA As Long = &H1
Public Const AC_SRC_OVER = &H0

Public Const UM_CLOSE_STARTMENU As Long = WM_USER + 1

Private m_GDIInitialized As Boolean

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Type ARGB
    r As Byte
    g As Byte
    b As Byte
    A As Byte
End Type

Public Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    rec As RECT
    hinst As Long
    lpszText As String
    lParam As Long
    lpReserved As Long
End Type

Public Const TTS_NOPREFIX = &H2
Public Const TTM_ADDTOOLA = (WM_USER + 4)
Public Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Public Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Public Const TTM_SETTITLE = (WM_USER + 32)
Public Const TTS_BALLOON = &H40
Public Const TTS_ALWAYSTIP = &H1
Public Const TTF_SUBCLASS = &H10
Public Const TTF_IDISHWND = &H1
Public Const TTM_SETDELAYTIME = (WM_USER + 3)

Public Const TTM_POP = (WM_USER + 28)

Public Function IsViPadInstalled() As Boolean
    IsViPadInstalled = FileExists(Environ$("appdata") & "\ViPad\ViPad.exe")
End Function

Public Function GenerateViPadAddToCommand(ByVal theTargetLinkFile As String)
    
    GenerateViPadAddToCommand = Environ$("appdata") & "\ViPad\ViPad.exe " & _
                                    """" & theTargetLinkFile & """"
End Function

Public Function LoadStringFromResource(theId As String, theType As String)

Dim theBinary() As Byte

    theBinary = LoadResData(theId, theType)
    LoadStringFromResource = StrConv(theBinary, vbUnicode)

End Function

Public Function TrimTrailingSlash(ByVal InputString As String) As String
    
    If Len(InputString) > 0 Then
        If Right$(InputString, 1) = "\" Then
            InputString = Left$(InputString, Len(InputString) - 1)
        End If
    End If
    
    TrimTrailingSlash = InputString
End Function

Public Function FindFormByName(ByVal szFormName As String) As Form

Dim thisForm As Form

    For Each thisForm In Forms
        If LCase$(thisForm.Name) = LCase$(szFormName) Then
            Set FindFormByName = thisForm
        End If
    Next
End Function

Public Function SendAppMessage(ByVal sourcehWnd As Long, ByVal destinationhWnd As Long, theData As String)

Dim tCDS As COPYDATASTRUCT
Dim dataToSend() As Byte

    dataToSend = theData

    With tCDS
        tCDS.lpData = VarPtr(dataToSend(0))
        tCDS.dwData = 87
        tCDS.cbData = UBound(dataToSend)
    End With
        
    SendMessage destinationhWnd, WM_COPYDATA, ByVal CLng(sourcehWnd), tCDS
End Function

Sub RemoveFromShellContextMenu(theType As String, Optional theText As String = CONTEXT_MENU)
    InitClasses_IfNeeded
    
    If Registry.RegObj.KeyExists(HKEY_CLASSES_ROOT, theType & "\shell\" & theText) Then
        Registry.RegObj.DeleteKey HKEY_CLASSES_ROOT, theType & "\shell\" & theText
    End If

End Sub

Sub AddToShellContextMenu(theType As String, Optional theText As String = CONTEXT_MENU)
    InitClasses_IfNeeded
    
Dim strKeyValue As String
    strKeyValue = App.Path & "\" & App.EXEName & ".exe" & " /pin " & """" & "%1" & """"

    If Not Registry.RegObj.KeyExists(HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command") Then
    
        Registry.RegObj.CreateKey HKEY_CLASSES_ROOT, theType & "\shell\" & theText
        Registry.RegObj.CreateKey HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command"
        
        Registry.RegObj.SetStringValue HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command", "", strKeyValue
    Else
        If Registry.RegObj.GetStringValue(HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command\", "", "") <> strKeyValue Then
            Registry.RegObj.SetStringValue HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command", "", strKeyValue
        End If
    End If

End Sub

Sub CreateFileAssociation(ByVal szExtension As String, ByVal szClassName As String, _
    ByVal szDescription As String, ByVal szExeProgram As String)
    
    ' ensure that there is a leading dot
    If Left$(szExtension, 1) <> "." Then
        szExtension = "." & szExtension
    End If
    
    Registry.RegObj.CreateKey HKEY_CLASSES_ROOT, szExtension
    Registry.RegObj.SetStringValue HKEY_CLASSES_ROOT, szExtension, vbNullString, szClassName
    Registry.RegObj.SetStringValue HKEY_CLASSES_ROOT, szClassName, "", szDescription
    
    Registry.RegObj.CreateKey HKEY_CLASSES_ROOT, szClassName & "\Shell\Open\Command"
    Registry.RegObj.SetStringValue HKEY_CLASSES_ROOT, szClassName & "\Shell\Open\Command", "", _
                        szExeProgram & " /install_theme " & " ""%1"""
End Sub

Public Function CheckBoxToBoolean(ByVal theValue As CheckBoxConstants) As Boolean

    If theValue = vbChecked Then
        CheckBoxToBoolean = True
    Else
        CheckBoxToBoolean = False
    End If

End Function

Public Function BooleanToCheckBox(ByVal theValue As Boolean) As CheckBoxConstants
    
    If theValue Then
        BooleanToCheckBox = vbChecked
    Else
        BooleanToCheckBox = vbUnchecked
    End If
    
End Function

Public Function ExtractXMLTextElement(ByRef parentElement As IXMLDOMElement, ByVal szElementName As String, ByVal DefaultValue As String) As String
    On Error GoTo Handler
    
    ExtractXMLTextElement = CStr(parentElement.selectSingleNode(szElementName).Text)
    Exit Function
Handler:
    ExtractXMLTextElement = DefaultValue
End Function

Public Function CreateXMLTextElement(ByRef sourceDoc As DOMDocument, ByRef parentElement As IXMLDOMElement, ByVal szElementName As String, ByVal szValue As String)
    
Dim element As IXMLDOMElement

    Set element = sourceDoc.createElement(szElementName)
    parentElement.appendChild element
    
    element.Text = szValue
End Function

Public Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left$(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim$(StrIn)
   End Select
End Function

Public Function ReconstructBackgroundImage(ByRef backgroundImage As GDIPImage, ByRef regionToDestroy As RECTL) As GDIPImage

Dim graphics As New GDIPGraphics
Dim newBackground As New GDIPBitmap
Dim regionLeft As gdiplus.RECTL
Dim encoder As New GDIPImageEncoderList


    newBackground.CreateFromSizeFormat backgroundImage.Width, backgroundImage.Height, GDIPlusWrapper.Format32bppArgb
    graphics.FromImage newBackground.Image
    
    regionLeft.Left = 0
    regionLeft.Top = 0
    regionLeft.Width = regionToDestroy.Left
    regionLeft.Height = backgroundImage.Height
    
    'graphics.DrawImageRectL backgroundImage, regionLeft
    
    'MsgBox regionToDestroy.Left + regionToDestroy.Right
    
    graphics.DrawImageRect backgroundImage, 0, 0, regionToDestroy.Left, backgroundImage.Height, 0, 0
    graphics.DrawImageRect backgroundImage, regionToDestroy.Left, 0, regionToDestroy.Right - regionToDestroy.Left, regionToDestroy.Top, regionToDestroy.Left, 0
    
    graphics.DrawImageRect backgroundImage, regionToDestroy.Right, 0, backgroundImage.Width, backgroundImage.Height, regionToDestroy.Right, 0
    graphics.DrawImageRect backgroundImage, regionToDestroy.Left, regionToDestroy.Bottom, regionToDestroy.Right - regionToDestroy.Left, backgroundImage.Height, regionToDestroy.Left, regionToDestroy.Bottom

    Set ReconstructBackgroundImage = newBackground.Image.Clone
    'newBackground.Image.Save "C:\b.png", encoder.EncoderForMimeType("image/png").CodecCLSID
End Function

Public Function MSHashString(Text As String) As Long
    HashData ByVal Text, Len(Text), MSHashString, Len(MSHashString)
End Function

Public Function IsInsideViComponent(X As Single, Y As Single, ByRef testComponent As GenericViElement, ByRef outClientPosition As POINTL) As Boolean

    If X > testComponent.Left And X < testComponent.Left + testComponent.Width And _
        Y > testComponent.Top And Y < testComponent.Top + testComponent.Height Then
        
        outClientPosition.X = X - testComponent.Left
        outClientPosition.Y = Y - testComponent.Top
        
        IsInsideViComponent = True
    End If

End Function

Public Sub Long2ARGB(ByVal LongARGB As Long, ByRef ARGB As ARGB)
    win.CopyMemory ARGB, LongARGB, 4
End Sub

Public Function isset(srcAny) As Boolean

    On Error GoTo Handler

Dim thisVarType As VbVarType: thisVarType = VarType(srcAny)

    If thisVarType = vbObject Then
        If Not srcAny Is Nothing Then
            isset = True
            Exit Function
        End If
    ElseIf thisVarType = vbArray Or _
           thisVarType = 8200 Then
           
            If UBound(srcAny) > 0 Then
                isset = True
                Exit Function
            End If
    Else
        isset = IsEmpty(srcAny)
        Exit Function
    End If

Handler:
    isset = False

End Function

Public Function RegisterAppRestart()
        
    If Not g_WindowsXP Then
        RegisterApplicationRestart vbNullString, ByVal 0
    End If
End Function

Public Function GetFileName(theFilePath As String) As String

Dim theDelim As String
    
    If InStr(theFilePath, "\") > 0 Then
        theDelim = "\"
    ElseIf InStr(theFilePath, "/") > 0 Then
        theDelim = "/"
    End If
    
    GetFileName = Right$(theFilePath, Len(theFilePath) - InStrRev(theFilePath, theDelim))
End Function

Public Function getAttribute_IgnoreError(ByRef theElement, attributeName As String, DefaultValue As Variant) As Variant
    On Error GoTo Handler
    
    getAttribute_IgnoreError = DefaultValue
    getAttribute_IgnoreError = theElement.getAttribute(attributeName)
    
    
    Exit Function
Handler:
    getAttribute_IgnoreError = DefaultValue
End Function

Public Function GetNativePath(strPath As String) As String

    If InStr(strPath, Environ$("systemdrive") & "\") > 0 Then
        strPath = Replace(strPath, Environ$("systemdrive") & "\", "")
    End If
    
    GetNativePath = Environ$("windir") & "\sysnative\..\..\" & strPath
    'End If

End Function

Public Function InitializeGDIIfNotInitialized() As Boolean
    If Not m_GDIInitialized Then
        ' Must call this before using any GDI+ call:
        If Not (GDIPlusCreate(True)) Then
            Exit Function
        End If
    
        m_GDIInitialized = True
    End If
    
    InitializeGDIIfNotInitialized = m_GDIInitialized
End Function

Public Function TopMost(lHWnd As Long)

    'typically called in the form load
    Call SetWindowPos(lHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Function

Function SetOwner(ByVal HwndtoUse, ByVal HwndofOwner) As Long

    SetOwner = SetWindowLong(HwndtoUse, GWL_HWNDPARENT, HwndofOwner)
End Function

Property Get CurrentDPI() As Long
    CurrentDPI = Registry.Read("HKCU\Control Panel\Desktop\WindowMetrics\AppliedDPI")
End Property

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
   MAKELPARAM = MakeLong(wLow, wHigh)
End Function

Public Function SetKeyboardActiveWindow(hWnd As Long) As Boolean

Dim ForegroundThreadID As Long
Dim ThisThreadID As Long

    SetKeyboardActiveWindow = False

    ForegroundThreadID = GetWindowThreadProcessId(GetForegroundWindow(), 0&)
    ThisThreadID = GetWindowThreadProcessId(hWnd, 0&)
    If AttachThreadInput(ThisThreadID, ForegroundThreadID, 1) = 1 Then
        BringWindowToTop hWnd
        SetForegroundWindow hWnd
        AttachThreadInput ThisThreadID, ForegroundThreadID, 0
        SetKeyboardActiveWindow = (GetForegroundWindow = hWnd)
    End If

End Function

Public Function TrackMouse(hWnd As Long) As Boolean

Dim ET As TrackMouseEvent

    TrackMouse = False
    
    'initialize structure
    ET.cbSize = Len(ET)
    ET.hwndTrack = hWnd
    ET.dwFlags = TME_LEAVE
    'start the tracking
    If Not TrackMouseEvent(ET) = 0 Then
        TrackMouse = True
    End If
    
End Function

Public Function CreateRect(Left, Top, Right, Bottom) As RECT

Dim r As RECT
    
    r.Left = Left
    r.Top = Top
    r.Right = Right
    r.Bottom = Bottom
    
    CreateRect = r

End Function

Public Function CreatePointF(Y As Single, X As Single) As POINTF

Dim p As POINTF
    p.X = X
    p.Y = Y
    
    CreatePointF = p
    
End Function

Public Function CreatePointL(Y As Long, X As Long) As POINTL

Dim p As POINTL
    p.X = X
    p.Y = Y
    
    CreatePointL = p
    
End Function

Public Sub RepaintWindow2( _
        ByRef objThis As Form, _
        Optional ByVal bClientAreaOnly As Boolean = True _
    )
    
Dim tR As RECT
Dim tP As POINTL

    If (bClientAreaOnly) Then
        GetClientRect objThis.hWnd, tR
    Else
        GetWindowRect objThis.hWnd, tR
        tP.X = tR.Left: tP.Y = tR.Top
        ScreenToClient objThis.hWnd, tP
        tR.Left = tP.X: tR.Top = tP.Y
        tP.X = tR.Right: tP.Y = tR.Bottom
        ScreenToClient objThis.hWnd, tP
        tR.Right = tP.X: tR.Bottom = tP.Y
    End If
    
    objThis.Height = objThis.Height - 15
    
    InvalidateRect objThis.hWnd, tR, 1
    UpdateWindow objThis.hWnd
    
    ReleaseDC objThis.hWnd, GetWindowDC(objThis.hWnd)
    RepaintWindow objThis.hWnd
    
    SendMessage ByVal objThis.hWnd, ByVal WM_PAINT, ByVal GetWindowDC(objThis.hWnd), ByVal 0

    objThis.Height = objThis.Height + 15
End Sub

Public Sub RepaintWindow(ByRef hWnd As Long)
    'verified it works
    
    If hWnd <> 0 Then
        Call RedrawWindow(hWnd, ByVal 0&, 0&, _
             RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
    End If
    
End Sub

Public Function ResourcesPath()
    ResourcesPath = g_resourcesPath
End Function

Public Function Is64bit() As Boolean
    Dim Handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    Handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If Handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

End Function

Public Function Wow64Wrapper(ByVal szPath As String)

    Wow64Wrapper = szPath

    If Is64bit Then
        If szPath <> vbNullString And Not FileExists(szPath) Then
            If FileExists(Replace(szPath, Environ$("ProgramFiles"), Environ$("ProgramW6432"))) Then
                Wow64Wrapper = Replace(szPath, Environ$("ProgramFiles"), Environ$("ProgramW6432"))
            ElseIf FileExists(Replace(LCase$(szPath), "system32", "sysnative")) Then
                Wow64Wrapper = Replace(LCase$(szPath), "system32", "sysnative")
            End If
        End If
    End If

End Function

Public Function ResolveLink(ByVal LnkPathName As String) As String
    On Error GoTo Handler

    If UCase$(Right$(LnkPathName, 3)) <> "LNK" Then
        ResolveLink = LnkPathName
        Exit Function
    End If
    
    LnkPathName = PathRemoveBlackSlash(LnkPathName)
    
Dim A As ShellLinkObject
Dim szLinkFileName As String

    Set A = GetShellLink(LnkPathName)
    ResolveLink = A.Target
    
    Exit Function
Handler:
    ResolveLink = LnkPathName
End Function

Public Function GetUserTilePath()

Dim emptyString As String
Dim picPath As String * 1024

    GetUserTilePathAPI ByVal emptyString, ByVal &H80000000, ByVal picPath, Len(picPath)
    GetUserTilePath = StrConv(picPath, vbFromUnicode)

End Function
