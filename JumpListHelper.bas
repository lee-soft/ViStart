Attribute VB_Name = "MRUHelper"
Option Explicit

Private Const EXPLORER_RECENTDOCS As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"

Private Const EXPLORER_OPENSAVEDOCS_XP As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSaveMRU"
Private Const EXPLORER_OPENSAVEDOCS_VISTA As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU"

Private EXPLORER_OPENSAVEDOCS As String

Function SetOpenSaveDocs()

Dim thisType As New RegistryKey

    thisType.RootKeyType = HKEY_CURRENT_USER
    EXPLORER_OPENSAVEDOCS = EXPLORER_OPENSAVEDOCS_XP
    
    thisType.Path = EXPLORER_OPENSAVEDOCS
    
    If thisType.GetLastError <> 0 Then
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

Dim thisKey As ViRegistryKey
Dim endFileNamePos As Long
Dim MRUList() As String
Dim lnkFileName As String
Dim MRUArrayIndex As Long

    If srcMRURoot Is Nothing Then Exit Function
    s_mruList = srcMRURoot.GetValueAsString("MRUList")
    
    

    If srcMRURoot.GetLastError = 0 Then
        While LenB(s_mruList) > 0
            thisMRU = MidB$(s_mruList, 1, 2)
            s_mruList = MidB$(s_mruList, LenB(thisMRU) + 1)
            
            If LenB(thisMRU) = 2 Then
            
                'Debug.Print "Reading:: " & CStr(thisMRU)
                
                thisMRUValue = srcMRURoot.GetValueAsString(CStr(thisMRU))
                lnkFileName = thisMRUValue
                
                Debug.Print "lnkFileName:: " & lnkFileName & "'"

                If FileExists(lnkFileName) Then
                    ReDim Preserve MRUList(MRUArrayIndex)
                    MRUList(MRUArrayIndex) = lnkFileName
                    
                    MRUArrayIndex = MRUArrayIndex + 1
                End If
            End If
        Wend
    Else
    
        s_mruList = srcMRURoot.GetValueAsString("MRUListEx")

        While LenB(s_mruList) > 0

            thisMRU = MidB$(s_mruList, 1, 4)
            s_mruList = MidB$(s_mruList, 5)
            
            thisMRU = GetDWord(CStr(thisMRU))

            If thisMRU > -1 Then
            
                thisMRUValue = srcMRURoot.GetValueAsString(CStr(thisMRU))
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
    End If
    
    GetMRUListForKey = MRUList
    
End Function

Public Function GetImageJumpList(ByVal srcImagePath As String)

    'Debug.Print "GetImageJumpList:: " & srcImagePath

    On Error GoTo Handler

Dim r_recentDocs As New ViRegistryKey
Dim r_openSaveDocs As New ViRegistryKey

Dim thisType As ViRegistryKey
Dim thisImagePath As String
Dim setJumpList As Boolean
Dim thisJumpList As New JumpList

    Set GetImageJumpList = thisJumpList
    srcImagePath = UCase$(StrEnd(srcImagePath, "\"))
    
    If Len(srcImagePath) = 0 Then Exit Function

    r_openSaveDocs.RootKeyType = HKEY_CURRENT_USER
    r_openSaveDocs.Path = EXPLORER_OPENSAVEDOCS

    r_recentDocs.RootKeyType = HKEY_CURRENT_USER
    r_recentDocs.Path = EXPLORER_RECENTDOCS

    thisJumpList.ImageName = srcImagePath
    srcImagePath = UCase$(srcImagePath)

    If isset(r_recentDocs.SubKeys) Then
        For Each thisType In r_recentDocs.SubKeys
            'thisImagePath = Ucase$(StrEnd(Trim$(GetEXEPathFromQuote(GetAbsolutePath(GetTypeHandlerPath(thisType.Name)))), "\"))

            Debug.Print thisType.Name & ":" & GetTypeHandlerPath(thisType.Name)
            
            If ExistInStringArray(GetTypeHandlersImageName(thisType.Name), srcImagePath) Then
                thisJumpList.AddMRURegKey thisType
                'GetMRUListForKey thisType
                
                setJumpList = True
            End If
        Next
    End If
    
    If isset(r_openSaveDocs.SubKeys) Then
        For Each thisType In r_openSaveDocs.SubKeys
            'thisImagePath = Ucase$(StrEnd(Trim$(GetEXEPathFromQuote(GetAbsolutePath(GetTypeHandlerPath("." & thisType.Name)))), "\"))
    
            If ExistInStringArray(GetTypeHandlersImageName(thisType.Name), srcImagePath) Then
            'If thisImagePath = srcImagePath Then
                thisJumpList.AddMRURegKey thisType
                setJumpList = True
            End If
        Next
    End If
    
    Exit Function
Handler:
    LogError Err.Description, "GetImageJumpList"
End Function

Public Function GetTypeHandlersImageName(srcType As String) As String()

Dim thisKey As New ViRegistryKey
Dim primaryCommand As String
Dim theChars() As Byte
Dim theCharIndex As Long
Dim returnHandlers() As String

    If Left$(srcType, 1) <> "." Then srcType = "." & srcType
    
    thisKey.RootKeyType = HKEY_CURRENT_USER
    thisKey.Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & srcType & "\OpenWithList"
    
    theChars = StrConv(thisKey.GetValueAsString("MRUList"), vbFromUnicode)
    For theCharIndex = LBound(theChars) To UBound(theChars)
        ReDim Preserve returnHandlers(theCharIndex)
        returnHandlers(theCharIndex) = thisKey.GetValueAsString(Chr$(theChars(theCharIndex)))
    Next
    
    GetTypeHandlersImageName = returnHandlers
    'typeFullName = thisKey.GetValueAsString()

End Function

Public Function GetTypeHandlerPath(srcType As String)

Dim thisKey As New ViRegistryKey
Dim typeFullName As String
Dim primaryCommand As String
    
    thisKey.RootKeyType = HKEY_CLASSES_ROOT
    thisKey.Path = srcType
    
    typeFullName = thisKey.GetValueAsString()
    
    thisKey.Path = typeFullName & "\shell"
    primaryCommand = thisKey.GetValueAsString()
    If primaryCommand = "" Then primaryCommand = "open"
    
    thisKey.Path = typeFullName & "\shell\" & primaryCommand & "\command"
    GetTypeHandlerPath = thisKey.GetValueAsString

    'Debug.Print srcType & "::" & GetTypeHandlerPath

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
