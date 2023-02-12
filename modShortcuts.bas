Attribute VB_Name = "AppLauncherHelper"
Option Explicit

Private Const GENERIC_READ As Long = &H80000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    lpOverlapped As Any) As Long

Private Type Buffer
    b(31) As Byte
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const SYNCHRONIZE = &H100000

Public Declare Function ShellExecuteW Lib "shell32.dll" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (Optional ByVal hOwner As Long, Optional ByVal UnknownP1 As Long, Optional ByVal UnknownP2 As Long, Optional ByVal szTitle As String, Optional ByVal szPrompt As String, Optional ByVal uFlags As Long) As Long

Private Declare Function ShellExecuteExW Lib "shell32.dll" (lpExecInfo As SHELLEXECUTEINFOW) As Long

Private Type SHELLEXECUTEINFOW
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As Long
    lpFile As Long
    lpParameters As Long
    lpDirectory As Long
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Const SE_ERR_NOASSOC As Long = 31
Private Const SE_ERR_NOTEXIST As Long = 2

Dim FSO As New FileSystemObject

Public Function SelectBestExecutionMethod(ByVal szFile As String)

Dim theFile As String
Dim thePathExtension As String

    theFile = LCase$(szFile)
    thePathExtension = PathFindExtension(szFile)
    
    If InStr(theFile, "cmd.exe /c") = 1 Then
        If Len(theFile) > 11 Then theFile = Mid$(theFile, 11)
        ShellCommand theFile
    ElseIf Left$(theFile, 8) = "explorer" Then
        
        On Error Resume Next
        Shell theFile
    Else
    
        'We can't trust the LINK file in Windows 8
        If g_Windows8 Or g_Windows81 Then
            ExplorerRun szFile
        Else
            If FolderExists(szFile) Then
                Shell "explorer.exe " & """" & szFile & """", vbNormalFocus
            Else
                'ExplorerRun was the old way, when it's not WindowsXP
                ShellExecuteW 0, StrPtr("Open"), StrPtr(szFile), 0, StrPtr(""), SW_SHOWNORMAL
            End If
        End If
        
        'CmdRun szFile, vbNormalFocus
    End If
        
End Function

Public Function CmdRun(sPath As String, Optional winStyle As VbAppWinStyle = vbHide)

    On Error GoTo Handler
    Shell "cmd.exe /c " & """" & sPath & """", winStyle

    Exit Function
Handler:

    CreateError "modEXE", "Cmd_Run", _
           "Parameter[1] = " & sPath & vbCrLf & _
           vbCrLf & _
           Err.Description
End Function

Public Function ExplorerRun(sPath As String, Optional sVisibility As VbAppWinStyle = vbNormalFocus)

    On Error GoTo Handler
    Shell "explorer.exe " & """" & sPath & """", sVisibility

    Exit Function
Handler:

    CreateError "modEXE", "Explorer_Run", _
           "Parameter[1] = " & sPath & vbCrLf & _
           vbCrLf & _
           Err.Description

End Function

Public Function ShellCommand(Program As String) As Boolean
    On Error GoTo Handler

    Shell "cmd.exe /c " & Program & " && exit", vbHide
    ShellCommand = True
    
    Exit Function
Handler:
    ShellCommand = False
    'g_WinError.ShowError "winStartBar", "Shell32", Err.description
    CreateError "", "ShellCommand", Err.Description

End Function

Public Function shell32(hWnd As Long, strPath As String) As Boolean
    
Dim lngResult As Long
    
    lngResult = ShellExecute(hWnd, 0, ByVal strPath, 0, 0, SW_SHOWNORMAL)
    shell32 = False
    
    If lngResult > 32 Then
        shell32 = True
    End If
End Function

Public Function ShellEx(ByVal strPath As String, Optional theVerb As String = vbNullString) As BOOL
    
Dim ShellExInfo As SHELLEXECUTEINFOW

Dim strEXE As String
Dim strParam As String

Dim lngFirstQ As Long
Dim lngSecondQ As Long

    On Error GoTo Handler

    If strPath = vbNullString Then
        ShellEx = APIFALSE
        Exit Function
    End If
    
    If InStr(strPath, """") Then
        lngFirstQ = InStr(strPath, """") + 1
        lngSecondQ = InStr(lngFirstQ, strPath, """")
        
        If lngFirstQ = 2 And (lngSecondQ < Len(strPath)) Then
            '"The.EXE" - Paramaters
            
            strParam = Mid$(strPath, lngSecondQ + 1)
            strEXE = Mid$(strPath, lngFirstQ, lngSecondQ - lngFirstQ)
        ElseIf lngFirstQ > 2 Then
            'The.EXE "The Parameters"
            
            strEXE = Mid$(strPath, 1, lngFirstQ - 2)
            strParam = Mid$(strPath, lngFirstQ - 1)
        End If
    Else
        'The.EXE
        strEXE = strPath
    End If
    
    ShellExInfo.lpFile = StrPtr(strEXE)
    ShellExInfo.lpParameters = StrPtr(strParam)
    ShellExInfo.lpVerb = StrPtr(theVerb)
    
    If LCase$(theVerb) = "properties" Then
        ShellExInfo.fMask = SEE_MASK_INVOKEIDLIST
    End If
    
    ShellExInfo.cbSize = Len(ShellExInfo)
    ShellExInfo.nShow = SW_SHOWNORMAL
    
    ShellEx = ShellExecuteExW(ShellExInfo)

    Exit Function
Handler:
    CreateError "AppLauncherHelper", "ShellEx", Err.Description
End Function

Function HexToStr(ByRef strHex)

Dim Length
Dim Max
Dim str
    
    Max = Len(strHex)
    For Length = 1 To Max Step 2
        str = str & Chr$("&h" & Mid$(strHex, Length, 2))
    Next
    HexToStr = str
    
End Function

Sub ShowRun(Optional ByRef lngHwnd As Long = 0)
    SHRunDialog lngHwnd
End Sub

Sub ShowFind()
    ShellExecuteW 0, StrPtr("find"), 0, _
      0, 0, vbNormalFocus
      
End Sub

Function StrToHex(s As Variant) As Variant
'
' Converts a string to a series of hexadecimal digits.
' For example, StrToHex(Chr$(9) & "A~") returns 09417E.
'
   Dim Temp As String, i As Integer
      If VarType(s) <> 8 Then
         StrToHex = s
      Else
         Temp = ""
      For i = 1 To Len(s)
         Temp = Temp & Format(Hex(Asc(Mid$(s, i, 1))), "00")
      Next i
         StrToHex = Temp
      End If
End Function

Function isFolderShortcut(ByVal sFilePath As String) As Boolean

Dim hFile As Long
Dim sBuffer As String * 32
    
    On Error GoTo Handler
    isFolderShortcut = False
    
    hFile = CreateFileW(StrPtr(sFilePath), _
            GENERIC_READ, _
            0, 0, _
            OPEN_EXISTING, _
            FILE_ATTRIBUTE_NORMAL, 0&)
             
    If hFile = -1 Then
    
        isFolderShortcut = True
        Exit Function
    End If
             
    ReadFile hFile, ByVal sBuffer, Len(sBuffer), 0, ByVal 0

    If Asc(Mid$(sBuffer, 21, 1)) = 139 Or Asc(Mid$(sBuffer, 21, 1)) = 131 Then
        isFolderShortcut = True
    End If
        
    CloseHandle hFile
    
    Exit Function
Handler:

    CloseHandle hFile
    isFolderShortcut = True

End Function

Public Function ProcessCount(ByVal theProcessName As String) As Long

On Error Resume Next

    Dim objWMIService, objProcess, colProcess
    Dim strComputer, strProcessKill
    strComputer = "."
    strProcessKill = "'" & theProcessName & ".exe" & "'"
    
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
    
    Set colProcess = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = " & strProcessKill)
    For Each objProcess In colProcess
        ProcessCount = ProcessCount + 1
    Next
    ' End of WMI Example of a Kill Process
End Function

Public Function KillProcess(ByVal theProcessName As String)

On Error Resume Next

    Dim objWMIService, objProcess, colProcess
    Dim strComputer, strProcessKill
    strComputer = "."
    strProcessKill = "'" & theProcessName & ".exe" & "'"
    
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
    
    Set colProcess = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = " & strProcessKill)
    For Each objProcess In colProcess
        objProcess.Terminate
    Next
    ' End of WMI Example of a Kill Process


End Function

