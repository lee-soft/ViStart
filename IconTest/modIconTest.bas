Attribute VB_Name = "modIconTest"
Option Explicit

'Isolated harness: find a recipe that extracts the REAL icon for every packaged
'(UWP/Store) app in shell:AppsFolder. Round 1 showed IShellItemImageFactory::
'GetImage(32, SIIGBF_ICONONLY) through a rebuilt shell item returns the generic
'document icon for true packaged apps (AUMIDs like Pkg_hash!App), while win32
'shortcut entries come out fine. This round runs a flag/size matrix per app and
'also probes fallback PIDL sources:
'  - SHParseDisplayName("::{CLSID}\AUMID")  (ViStart's current fallback - HR recorded)
'  - Folder.ParseName(AUMID) automation     (candidate replacement fallback)
'
'Output: apps.csv (one row per app), flags.csv (one row per app per recipe),
'PNG dumps of every recipe for true-UWP entries.
'
'Usage: IconTest.exe <output-folder>   (defaults to <exe folder>\dump)

Private Const APPSFOLDER_CLSID As String = "{4234d49b-0245-4df3-b780-3893943456e1}"
Private Const IID_ISHELLITEMIMAGEFACTORY As String = "{BCC18B79-BA16-442F-80C4-8A59C30C463B}"
Private Const SHELL_APPSFOLDER_PREFIX As String = "SHELL:APPSFOLDER\"

Private Const CC_STDCALL As Long = 4

Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const CLSID_PNG As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"

Private Type BITMAPSTRUCT
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO32
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 3) As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EXTRACTRESULT
    ok As Boolean
    stage As String        'bind / getimage / bits / ok
    hr As Long
    bmpW As Long
    bmpH As Long
    bmpBpp As Long
    alphaPct As Long
    overCount As Long
    hash As Double
End Type

Private Type RECIPE
    label As String
    size As Long
    flags As Long
End Type

Private Declare Function SHGetIDListFromObject Lib "shell32" (ByVal punk As IUnknown, ppidl As Long) As Long
Private Declare Function SHCreateItemFromIDList Lib "shell32" (ByVal pidl As Long, riid As Any, ppv As Long) As Long
Private Declare Function SHParseDisplayName Lib "shell32" (ByVal pszName As Long, ByVal pbc As Long, ppidl As Long, ByVal sfgaoIn As Long, psfgaoOut As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBmp As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, lpvBits As Any, lpbi As BITMAPINFO32, ByVal uUsage As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, scan0 As Any, bitmap As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal fileName As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long

Private m_gdipToken As Long

Public Sub Main()

Dim outDir As String
Dim shellApp As Object
Dim appsFolder As Object
Dim thisItem As Object
Dim parsedItem As Object
Dim seen As New Collection
Dim theName As String
Dim thePath As String
Dim aumid As String
Dim thisPidl As Long
Dim parsePidl As Long
Dim parseName As String
Dim dummySfgao As Long
Dim res As EXTRACTRESULT
Dim resPN As EXTRACTRESULT
Dim bits() As Byte
Dim idx As Long
Dim nWin32 As Long
Dim isUwp As Boolean
Dim parseHR As Long
Dim pnState As String
Dim baseHash As Double
Dim gdipInput As GdiplusStartupInput
Dim recipes() As RECIPE
Dim r As Long

    outDir = Trim$(Replace(Command$, """", vbNullString))
    If Len(outDir) = 0 Then outDir = App.Path & "\dump"
    On Error Resume Next
    MkDir outDir
    On Error GoTo 0

    gdipInput.GdiplusVersion = 1
    If GdiplusStartup(m_gdipToken, gdipInput, 0) <> 0 Then Exit Sub

    recipes = BuildRecipes()

    Open outDir & "\apps.csv" For Output As #1
    Print #1, "idx,name,path,isUwp,enumPidl,parseHR,parseNameState,parseNameHashSame,parseNameUpperState,parseNameUpperHashSame"
    Open outDir & "\flags.csv" For Output As #2
    Print #2, "idx,recipe,stage,hr,w,h,bpp,alphaPct,over,hash"

    Set shellApp = CreateObject("Shell.Application")
    Set appsFolder = shellApp.Namespace("shell:::" & APPSFOLDER_CLSID)
    If appsFolder Is Nothing Then GoTo Cleanup

    For Each thisItem In appsFolder.Items
        theName = vbNullString
        thePath = vbNullString

        On Error Resume Next
        theName = thisItem.Name
        thePath = thisItem.Path
        Err.Clear
        On Error GoTo 0

        If Len(theName) = 0 Or Len(thePath) = 0 Then GoTo NextItem

        If GetFileAttributesW(StrPtr(thePath)) <> -1 Then
            nWin32 = nWin32 + 1
            GoTo NextItem
        End If

        On Error Resume Next
        seen.Add 1, UCase$(thePath)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextItem
        On Error GoTo 0

        idx = idx + 1

        'True packaged apps have PackageFamilyName!AppId paths - no slashes
        isUwp = (InStr(thePath, "!") > 0 And InStr(thePath, "\") = 0 And InStr(thePath, "/") = 0)

        aumid = thePath
        If UCase$(Left$(aumid, Len(SHELL_APPSFOLDER_PREFIX))) = SHELL_APPSFOLDER_PREFIX Then
            aumid = Mid$(aumid, Len(SHELL_APPSFOLDER_PREFIX) + 1)
        End If

        'PIDL off the live FolderItem, run the whole recipe matrix against it
        thisPidl = 0
        baseHash = 0
        If SHGetIDListFromObject(thisItem, thisPidl) = 0 And thisPidl <> 0 Then
            For r = 0 To UBound(recipes)
                ZeroResult res
                Erase bits
                ExtractFromPidl thisPidl, recipes(r).size, recipes(r).flags, res, bits
                If r = 0 And res.ok Then baseHash = res.hash
                Print #2, idx & "," & recipes(r).label & "," & ResultCsv(res)
                If res.ok And isUwp Then
                    SavePng bits, res.bmpW, res.bmpH, outDir & "\" & Format$(idx, "000") & "_" & recipes(r).label & ".png"
                End If
            Next
            CoTaskMemFree thisPidl
        End If

        'ViStart's current fallback: SHParseDisplayName on the CLSID-rooted name
        parseName = "::" & APPSFOLDER_CLSID & "\" & aumid
        parsePidl = 0
        parseHR = SHParseDisplayName(StrPtr(parseName), 0, parsePidl, 0, dummySfgao)
        If parsePidl <> 0 Then CoTaskMemFree parsePidl

        'Candidate replacement fallback: Folder.ParseName automation, tried with
        'the original-case AUMID and the uppercased one ViStart's icon cache uses
        Dim pnUpperState As String
        Dim resPNU As EXTRACTRESULT
        pnState = TryParseName(appsFolder, aumid, recipes(0), resPN)
        pnUpperState = TryParseName(appsFolder, UCase$(aumid), recipes(0), resPNU)

        Print #1, idx & "," & Csv(theName) & "," & Csv(thePath) & "," & IIf(isUwp, 1, 0) & "," & _
            IIf(thisPidl <> 0, 1, 0) & "," & HrHex(parseHR) & "," & pnState & "," & _
            IIf(resPN.ok And baseHash <> 0, IIf(resPN.hash = baseHash, "same", "DIFF"), "n/a") & "," & _
            pnUpperState & "," & _
            IIf(resPNU.ok And baseHash <> 0, IIf(resPNU.hash = baseHash, "same", "DIFF"), "n/a")

NextItem:
    Next

    Print #1, ""
    Print #1, "SUMMARY,packaged=" & idx & ",win32skipped=" & nWin32

Cleanup:
    Close #2
    Close #1
    GdiplusShutdown m_gdipToken

End Sub

'Folder.ParseName(AUMID) -> PIDL -> extract with the given recipe. Returns a
'state string for the CSV and fills res on success.
Private Function TryParseName(ByVal appsFolder As Object, ByVal aumid As String, ByRef r As RECIPE, ByRef res As EXTRACTRESULT) As String

Dim parsedItem As Object
Dim parsePidl As Long
Dim bits() As Byte

    ZeroResult res
    TryParseName = "nothing"

    On Error Resume Next
    Set parsedItem = appsFolder.ParseName(aumid)
    Err.Clear
    On Error GoTo 0
    If parsedItem Is Nothing Then Exit Function

    parsePidl = 0
    If SHGetIDListFromObject(parsedItem, parsePidl) <> 0 Or parsePidl = 0 Then
        TryParseName = "nopidl"
        Exit Function
    End If

    ExtractFromPidl parsePidl, r.size, r.flags, res, bits
    CoTaskMemFree parsePidl

    If res.ok Then
        TryParseName = "ok"
    Else
        TryParseName = res.stage & ":" & HrHex(res.hr)
    End If

End Function

'Recipe 0 is the baseline (what ViStart ships today)
Private Function BuildRecipes() As RECIPE()

Dim list(0 To 9) As RECIPE

    SetRecipe list(0), "32f4", 32, &H4      'ICONONLY (current ViStart)
    SetRecipe list(1), "32f0", 32, &H0      'default thumbnail/RESIZETOFIT
    SetRecipe list(2), "32f1", 32, &H1      'BIGGERSIZEOK
    SetRecipe list(3), "32f80", 32, &H80    'ICONBACKGROUND
    SetRecipe list(4), "32f84", 32, &H84    'ICONONLY + ICONBACKGROUND
    SetRecipe list(5), "32f104", 32, &H104  'ICONONLY + SCALEUP
    SetRecipe list(6), "48f4", 48, &H4
    SetRecipe list(7), "96f0", 96, &H0
    SetRecipe list(8), "256f0", 256, &H0
    SetRecipe list(9), "16f0", 16, &H0      'ViStart's small-icon size

    BuildRecipes = list

End Function

Private Sub SetRecipe(ByRef r As RECIPE, ByVal label As String, ByVal size As Long, ByVal flags As Long)
    r.label = label
    r.size = size
    r.flags = flags
End Sub

Private Sub ZeroResult(ByRef res As EXTRACTRESULT)
    Dim blank As EXTRACTRESULT
    res = blank
    res.stage = "?"
End Sub

Private Sub ExtractFromPidl(ByVal thisPidl As Long, ByVal iconSize As Long, ByVal siigbf As Long, ByRef res As EXTRACTRESULT, ByRef outBits() As Byte)

Dim iidFactory(0 To 3) As Long
Dim pFactory As Long
Dim hBmp As Long

    CLSIDFromString StrPtr(IID_ISHELLITEMIMAGEFACTORY), iidFactory(0)

    res.stage = "bind"
    res.hr = SHCreateItemFromIDList(thisPidl, iidFactory(0), pFactory)
    If res.hr <> 0 Or pFactory = 0 Then Exit Sub

    res.stage = "getimage"
    res.hr = CallGetImage(pFactory, iconSize, siigbf, hBmp)
    ComRelease pFactory
    If res.hr <> 0 Or hBmp = 0 Then Exit Sub

    res.stage = "bits"
    If AnalyzeBitmap(hBmp, res, outBits) Then
        res.stage = "ok"
        res.ok = True
    End If
    DeleteObject hBmp

End Sub

'IShellItemImageFactory::GetImage is vtable slot 3 (offset 12). SIZE is passed
'by value = two dword slots on x86 stdcall. Returns the method HRESULT.
Private Function CallGetImage(ByVal pFactory As Long, ByVal iconSize As Long, ByVal siigbf As Long, ByRef hBmp As Long) As Long

Dim hr As Long
Dim vParams(0 To 3) As Variant
Dim vTypes(0 To 3) As Integer
Dim vPtrs(0 To 3) As Long
Dim vReturn As Variant
Dim i As Long

    hBmp = 0
    vParams(0) = iconSize          'SIZE.cx
    vParams(1) = iconSize          'SIZE.cy
    vParams(2) = siigbf
    vParams(3) = VarPtr(hBmp)

    For i = 0 To 3
        vTypes(i) = vbLong
        vPtrs(i) = VarPtr(vParams(i))
    Next

    hr = DispCallFunc(pFactory, 12&, CC_STDCALL, vbLong, 4&, vTypes(0), vPtrs(0), vReturn)
    If hr <> 0 Then
        CallGetImage = hr
    Else
        CallGetImage = CLng(vReturn)
    End If

End Function

'IUnknown::Release is vtable slot 2 (offset 8)
Private Sub ComRelease(ByVal pUnknown As Long)

Dim vReturn As Variant

    DispCallFunc pUnknown, 8&, CC_STDCALL, vbLong, 0&, ByVal 0&, ByVal 0&, vReturn

End Sub

'Pull the pixels out as top-down 32bpp BGRA and gather stats
Private Function AnalyzeBitmap(ByVal hBmp As Long, ByRef res As EXTRACTRESULT, ByRef outBits() As Byte) As Boolean

Dim bm As BITMAPSTRUCT
Dim bi As BITMAPINFO32
Dim hdc As Long
Dim i As Long
Dim nPixels As Long
Dim nAlpha As Long
Dim a As Long

    If GetObjectAPI(hBmp, LenB(bm), bm) = 0 Then Exit Function

    res.bmpW = bm.bmWidth
    res.bmpH = bm.bmHeight
    res.bmpBpp = bm.bmBitsPixel

    If bm.bmWidth <= 0 Or bm.bmHeight <= 0 Then Exit Function

    bi.bmiHeader.biSize = 40
    bi.bmiHeader.biWidth = bm.bmWidth
    bi.bmiHeader.biHeight = -bm.bmHeight   'top-down
    bi.bmiHeader.biPlanes = 1
    bi.bmiHeader.biBitCount = 32
    bi.bmiHeader.biCompression = 0

    ReDim outBits(0 To bm.bmWidth * bm.bmHeight * 4 - 1)

    hdc = GetDC(0)
    i = GetDIBits(hdc, hBmp, 0, bm.bmHeight, outBits(0), bi, 0)
    ReleaseDC 0, hdc
    If i <> bm.bmHeight Then Exit Function

    nPixels = bm.bmWidth * bm.bmHeight
    For i = 3 To UBound(outBits) Step 4
        a = outBits(i)
        If a > 0 Then nAlpha = nAlpha + 1
        If outBits(i - 3) > a Or outBits(i - 2) > a Or outBits(i - 1) > a Then
            res.overCount = res.overCount + 1
        End If
        res.hash = res.hash + CDbl(outBits(i - 3)) * (1 + ((i - 3) Mod 251)) _
                            + CDbl(outBits(i - 2)) * (1 + ((i - 2) Mod 251)) _
                            + CDbl(outBits(i - 1)) * (1 + ((i - 1) Mod 251)) _
                            + CDbl(a) * (1 + (i Mod 251))
    Next

    res.alphaPct = (nAlpha * 100) \ nPixels
    AnalyzeBitmap = True

End Function

'Write the top-down BGRA buffer out as a real PNG so alpha survives inspection
Private Sub SavePng(bits() As Byte, ByVal w As Long, ByVal h As Long, ByVal szFile As String)

Dim clsidEnc(0 To 3) As Long
Dim pBmp As Long

    If w <= 0 Or h <= 0 Then Exit Sub
    CLSIDFromString StrPtr(CLSID_PNG), clsidEnc(0)

    If GdipCreateBitmapFromScan0(w, h, w * 4, PixelFormat32bppARGB, bits(0), pBmp) = 0 Then
        GdipSaveImageToFile pBmp, StrPtr(szFile), clsidEnc(0), ByVal 0&
        GdipDisposeImage pBmp
    End If

End Sub

Private Function ResultCsv(ByRef res As EXTRACTRESULT) As String
    ResultCsv = res.stage & "," & HrHex(res.hr) & "," & res.bmpW & "," & res.bmpH & _
        "," & res.bmpBpp & "," & res.alphaPct & "," & res.overCount & "," & CStr(res.hash)
End Function

Private Function HrHex(ByVal hr As Long) As String
    If hr = 0 Then
        HrHex = "0"
    Else
        HrHex = "0x" & Right$("00000000" & Hex$(hr), 8)
    End If
End Function

Private Function Csv(ByVal s As String) As String
    Csv = """" & Replace(s, """", "'") & """"
End Function
