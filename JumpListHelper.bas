Attribute VB_Name = "MRUHelper"
Option Explicit

Private Const EXPLORER_RECENTDOCS As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"

Private Const EXPLORER_OPENSAVEDOCS_XP As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSaveMRU"
Private Const EXPLORER_OPENSAVEDOCS_VISTA As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU"

Private EXPLORER_OPENSAVEDOCS As String

Function SetOpenSaveDocs()
    On Error GoTo RegistryError
    
    EXPLORER_OPENSAVEDOCS = EXPLORER_OPENSAVEDOCS_XP
    If Registry.CurrentUser.OpenSubKey(EXPLORER_OPENSAVEDOCS) Is Nothing Then
        EXPLORER_OPENSAVEDOCS = EXPLORER_OPENSAVEDOCS_VISTA
    End If
End Function

Public Function GetMRUListForKey(ByRef srcMRURoot As RegistryKey) As String()
    On Error Resume Next
    
    Debug.Print srcMRURoot.Path

Dim s_mruList As String
Dim thisMRU
Dim thisMRUValue As String
Dim thisLnkName As String

Dim endFileNamePos As Long
Dim MRUList() As String
Dim lnkFileName As String
Dim MRUArrayIndex As Long

    If srcMRURoot Is Nothing Then Exit Function
    s_mruList = srcMRURoot.GetValue("MRUList")
    
    While LenB(s_mruList) > 0
        thisMRU = MidB$(s_mruList, 1, 2)
        s_mruList = MidB$(s_mruList, LenB(thisMRU) + 1)
        
        If LenB(thisMRU) = 2 Then
        
            'Debug.Print "Reading:: " & CStr(thisMRU)
            
            thisMRUValue = srcMRURoot.GetValue(CStr(thisMRU))
            lnkFileName = thisMRUValue
            
            Debug.Print "lnkFileName:: " & lnkFileName & "'"

            If FileExists(lnkFileName) Then
                ReDim Preserve MRUList(MRUArrayIndex)
                MRUList(MRUArrayIndex) = lnkFileName
                
                MRUArrayIndex = MRUArrayIndex + 1
            End If
        End If
    Wend
    
    s_mruList = srcMRURoot.GetValue("MRUListEx")

    While LenB(s_mruList) > 0

        thisMRU = MidB$(s_mruList, 1, 4)
        s_mruList = MidB$(s_mruList, 5)
        
        thisMRU = GetDWord(CStr(thisMRU))

        If thisMRU > -1 Then
        
            thisMRUValue = srcMRURoot.GetValue(CStr(thisMRU))
            endFileNamePos = 1
            
            'Debug.Print thisMRUValue
            
            'Chr$(0) is actually a double byte ZERO ChrB(0) is a single byte
            'Remember strings are double-byte in VB6
            endFileNamePos = InStrB(thisMRUValue, Chr$(0))

            If endFileNamePos > 1 Then
                thisLnkName = MidB$(thisMRUValue, 1, endFileNamePos)
                
                If Len(thisLnkName) > 3 Then
                
                    If Not (Right$(thisLnkName, 4) = ".lnk") And InStr(thisLnkName, ".") > 0 Then
                        
                        'Debug.Print "thisLnkName:: " & thisLnkName
                        
                        If FileExists(Environ$("userprofile") & "\Recent\" & Left$(thisLnkName, InStrRev(thisLnkName, ".") - 1) & ".lnk") Then
                            thisLnkName = Left$(thisLnkName, InStrRev(thisLnkName, ".") - 1) & ".lnk"
                        Else
                            thisLnkName = thisLnkName & ".lnk"
                        End If
                    End If
                End If
                
                lnkFileName = ResolveLink(Environ$("userprofile") & "\Recent\" & thisLnkName)
                
                If FileExists(lnkFileName) Then
                    ReDim Preserve MRUList(MRUArrayIndex)
                    MRUList(MRUArrayIndex) = lnkFileName
                    
                    MRUArrayIndex = MRUArrayIndex + 1
                End If
            End If
        End If
    Wend

    Err.Clear
    GetMRUListForKey = MRUList
    
Handler:
    Err.Clear
    LogError "GetMRUListForKey", ""
End Function

Public Function GetImageJumpList(ByVal srcImagePath As String) As JumpList
    On Error GoTo Handler

Dim r_recentDocs As RegistryKey
Dim r_openSaveDocs As RegistryKey

Dim thisTypeNameColItem As Variant
Dim thisTypeName As String

Dim thisImagePath As String
Dim setJumpList As Boolean
Dim thisJumpList As New JumpList

    Set GetImageJumpList = thisJumpList
    srcImagePath = UCase$(StrEnd(srcImagePath, "\"))
    
    If Len(srcImagePath) = 0 Then Exit Function
    
    Set r_openSaveDocs = Registry.CurrentUser.OpenSubKey(EXPLORER_OPENSAVEDOCS)
    Set r_recentDocs = Registry.CurrentUser.OpenSubKey(EXPLORER_RECENTDOCS)

    thisJumpList.ImageName = srcImagePath
    srcImagePath = UCase$(srcImagePath)

    For Each thisTypeNameColItem In r_recentDocs.GetSubKeyNames
        thisTypeName = CStr(thisTypeNameColItem)

        If ExistInStringArray(GetTypeHandlersImageName(thisTypeName), srcImagePath) Then
            thisJumpList.AddMRURegKey r_recentDocs.OpenSubKey(thisTypeName)
            setJumpList = True
        End If
    Next
    
    For Each thisTypeNameColItem In r_openSaveDocs.GetSubKeyNames
        thisTypeName = CStr(thisTypeNameColItem)

        If ExistInStringArray(GetTypeHandlersImageName(thisTypeName), srcImagePath) Then
            thisJumpList.AddMRURegKey r_openSaveDocs.OpenSubKey(thisTypeName)
            setJumpList = True
        End If
    Next
    
    Exit Function
Handler:
    LogError Err.Description, "GetImageJumpList"
End Function

Public Function GetTypeHandlersImageName(srcType As String) As String()

Dim thisKey As RegistryKey
Dim primaryCommand As String
Dim theChars() As Byte
Dim theCharIndex As Long
Dim returnHandlers() As String

    If Left$(srcType, 1) <> "." Then srcType = "." & srcType
    Set thisKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & srcType & "\OpenWithList")
    If thisKey Is Nothing Then
        Exit Function
    End If
    
    theChars = StrConv(thisKey.GetValue("MRUList"), vbFromUnicode)
    For theCharIndex = LBound(theChars) To UBound(theChars)
        ReDim Preserve returnHandlers(theCharIndex)
        returnHandlers(theCharIndex) = thisKey.GetValue(Chr$(theChars(theCharIndex)))
    Next
    
    GetTypeHandlersImageName = returnHandlers
    Exit Function
End Function

Public Function GetTypeHandlerPath(ByVal srcType As String) As String

    
Dim thisKey As RegistryKey
Dim typeFullName As String
Dim primaryCommand As String
    
    On Error GoTo HandleInvalidSubKey
    
    Set thisKey = Registry.ClassesRoot.OpenSubKey(srcType)
    typeFullName = thisKey.GetValue("")
    
    Set thisKey = Registry.ClassesRoot.OpenSubKey(typeFullName & "\shell")
    primaryCommand = thisKey.GetValue("")
    
    If primaryCommand = "" Then primaryCommand = "open"
    
    Set thisKey = Registry.ClassesRoot.OpenSubKey(typeFullName & "\shell\" & primaryCommand & "\command")
    GetTypeHandlerPath = thisKey.GetValue("")
    
    Exit Function
HandleInvalidSubKey:
End Function

Public Function GetEXEPathFromQuote(ByVal srcPath As String)
    On Error GoTo Handler
    
Dim A As Long
Dim b As Long
Dim spliceA As String
Dim spliceB As String
Dim ret As String

    A = InStr(srcPath, """") + 1
    b = InStr(A, srcPath, """")
    
    If (A <> 2) Then
        If (A > 1) Then
            'would fetch path in this situation:  C:\blabla\notepad.exe "%1"
            GetEXEPathFromQuote = Trim$(Mid$(srcPath, 1, A - 2))
            Exit Function
        Else
            'would fetch path in this situation:  C:\blabla\notepad.exe %1
            A = InStr(srcPath, "%") - 1
            If A > 0 Then
                GetEXEPathFromQuote = Left$(srcPath, A)
            Else
                GetEXEPathFromQuote = srcPath
            End If
            
            Exit Function
        End If
    End If
    
    If (A > 1 And b > 0 And _
        b > A) Then
        
        GetEXEPathFromQuote = Mid$(srcPath, A, (b - A))
        Exit Function
    Else
        A = InStr(srcPath, "%") - 1
        If (A > 0) Then
            GetEXEPathFromQuote = Mid$(srcPath, 1, A)
            Exit Function
        End If
    End If
    
    Exit Function
Handler:
    
    GetEXEPathFromQuote = srcPath
End Function
'Replaces all enviromental variables with their absolute equivalents
'It doesn't require that a path be valid either
Public Function GetAbsolutePath(ByVal srcPath As String)
    
Dim A As Long
Dim b As Long
Dim varName As String
Dim spliceA As String
Dim spliceB As String
Dim ret As String

    A = InStr(srcPath, "%") + 1
    b = InStr(A, srcPath, "%")
    
    If (A > 1 And b > 0 And _
        b > A) Then
        
        varName = Mid$(srcPath, A, (b - A))
        
        spliceA = Mid$(srcPath, 1, A - 2)
        spliceB = Mid$(srcPath, b + 1)
        
        ret = spliceA & Environ$(varName) & spliceB
    Else
        GetAbsolutePath = srcPath
        Exit Function
    End If
    
    If InStr(ret, "%") > 0 Then
        GetAbsolutePath = GetAbsolutePath(ret)
    Else
        GetAbsolutePath = ret
    End If
    
End Function
