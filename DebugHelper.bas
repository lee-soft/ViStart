Attribute VB_Name = "DebugHelper"
Option Explicit

Public g_tsTextFile As TextStream

Public Function GetWindowClassNameByHwnd(ByVal hWnd As Long) As String

Dim lReturn As Long
Dim mvarCaption As String

    mvarCaption = Space$(256)
    lReturn = GetClassName(hWnd, mvarCaption, Len(mvarCaption))
    
    If lReturn Then
        GetWindowClassNameByHwnd = Left$(mvarCaption, lReturn)
    End If

End Function

Public Function GetWindowNameByHwnd(ByVal hWnd As Long) As String

Dim lReturn As Long
Dim mvarCaption As String

    mvarCaption = Space$(256)
    lReturn = GetWindowText(hWnd, mvarCaption, Len(mvarCaption))
    
    If lReturn Then
        GetWindowNameByHwnd = Left$(mvarCaption, lReturn)
    End If

End Function


Public Function PrintBinary(theString As String)

Dim outputString As String
Dim theBinary() As Byte
Dim theBinaryIndex As Long

    theBinary = theString
    
    For theBinaryIndex = LBound(theBinary) To UBound(theBinary)
        outputString = outputString & "[" & theBinary(theBinaryIndex) & "] "
    Next
    
    Debug.Print outputString

End Function

Public Function CreateError(strModule As String, strFunction As String, strReason As String)
    LogError strReason, strModule & "::" & strFunction
End Function

Public Function WriteLine(strLine As String)
    g_tsTextFile.WriteLine "[" & Now & "]-" & strLine
End Function

Public Sub LogError(ByVal sDesc As String, Optional ByVal sFrom As String = "General")
  
    Debug.Print "APP ERROR; " & sDesc & " - " & sFrom
    
    Dim FileNum As Integer

    FileNum = FreeFile
    
    Open Environ$("appdata") & "\ViStart\errors.log" For Append As FileNum
        Write #FileNum, sDesc, sFrom, Now()
    Close FileNum
End Sub

