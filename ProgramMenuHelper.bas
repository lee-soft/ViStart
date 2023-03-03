Attribute VB_Name = "ProgramMenuHelper"
Option Explicit
'Various functions that minipulate Target collections

Const cNodeHeaderBoundary As Integer = 4

Private Declare Function VerQueryValueA Lib "Version.dll" _
         (pBlock As Any, _
         ByVal lpSubBlock As String, _
         lplpBuffer As Any, _
         puLen As Long) As Long
         
 Private Declare Function lstrcpyA Lib "kernel32" _
         (ByVal lpString1 As String, _
         ByVal lpString2 As Long) As Long

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        m_logger = LogManager.GetLogger("ProgramMenuHelper")
    End If
    
    Set Logger = m_logger
End Property

Function CreateProgramFromPath(ByVal szProgramPath As String) As clsProgram

Dim thisProgram As clsProgram

    Set thisProgram = New clsProgram
    Set CreateProgramFromPath = thisProgram
    
    Set thisProgram.Icon = IconManager.GetViIcon(szProgramPath, True)
    thisProgram.Caption = ExtOrNot(GetFileName(szProgramPath))
    thisProgram.Path = szProgramPath
End Function

Function CreateProgramFromNode(ByVal sourceNode As INode) As clsProgram

Dim thisProgram As clsProgram

    Set thisProgram = New clsProgram
    Set CreateProgramFromNode = thisProgram
    
    thisProgram.szIcon = sourceNode.Icon.IconPath
    Set thisProgram.Icon = IconManager.GetViIcon(thisProgram.szIcon, True)
    thisProgram.Caption = sourceNode.Caption
    thisProgram.Path = sourceNode.Tag
        
End Function

    
Function GetAppDescription(ByVal aExeName As String) As String
    On Error GoTo Handler

Dim lBufferLen As Long
Dim bBuffer() As Byte

Dim lDummy As Long
Dim lReceive  As Long
Dim pRecieve As Long
Dim sBuffer As String

    lBufferLen = GetFileVersionInfoSize(aExeName, lDummy)
    ReDim bBuffer(lBufferLen)
     
    GetFileVersionInfo aExeName, 0&, lBufferLen, bBuffer(0)
    
    'VerQueryValue bBuffer(0), "\StringFileInfo\040904B0\FileDescription", pRecieve, lReceive ' 040904E4 (Crashes randomy in XP and maybe more)
    VerQueryValueA bBuffer(0), "\StringFileInfo\040904B0\FileDescription", pRecieve, lReceive ' 040904E4
    
    sBuffer = String$(255, 0)
    lstrcpyA sBuffer, pRecieve
    
    sBuffer = Mid$(sBuffer, 1, InStr(sBuffer, Chr$(0)) - 1)
    GetAppDescription = sBuffer

    Exit Function
Handler:
    Logger.Error Err.Description, "GetAppDescription", aExeName
End Function

Public Function GetStringFromPointer(ByVal PtrStr As Long)

Dim sBuffer As String

    sBuffer = String$(255, 0)
    lstrcpyA sBuffer, PtrStr
    
    sBuffer = Mid$(sBuffer, 1, InStr(sBuffer, Chr$(0)) - 1)
    GetStringFromPointer = sBuffer

End Function

Public Function GetString(ByVal PtrStr As Long) As String
   Dim StrBuff As String * 256
   'Check for zero address
   If PtrStr = 0 Then
      GetString = vbNullString
      Exit Function
   End If
   'Copy data from PtrStr to buffer.
   win.CopyMemory ByVal StrBuff, ByVal PtrStr, 256
   'Strip any trailing nulls from string.
   GetString = StripNulls(StrBuff)
End Function

Public Function StripNulls(OriginalStr As String) As String
   'Strip any trailing nulls from input string.
   If (InStr(OriginalStr, Chr$(0)) > 0) Then
      OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr$(0)) - 1)
   End If
  'Return modified string.
   StripNulls = OriginalStr
End Function

Public Function IsValidArray(arr As Variant) As Boolean
    On Error Resume Next
    If LBound(arr) = UBound(arr) Then
        IsValidArray = False
        Exit Function
    Else
        IsValidArray = True
    End If
End Function

Public Function MakeSearchable(ByVal srcString As String) As String

Dim lngCharPosition As Long

    lngCharPosition = InStrRev(srcString, ".")
    If lngCharPosition > 1 Then
        srcString = Mid$(srcString, 1, lngCharPosition - 1) & " " & Mid$(srcString, lngCharPosition)
    End If
    
    MakeSearchable = UCase$(srcString)

End Function

Function NodeSize(ByRef cTarget) As Integer

Dim Obj As Object
    NodeSize = cTarget.count - (cNodeHeaderBoundary - 1)

End Function

Function IsCollection(ByRef cTarget) As Boolean

On Error GoTo Handler

    If cTarget.count <> 0 Then
        IsCollection = True
    End If
    
    Exit Function
Handler:
    IsCollection = False

End Function

Function ExistInCol(ByRef cTarget As Collection, sKey) As Boolean

    On Error GoTo Handler
    ExistInCol = Not (IsEmpty(cTarget(sKey)))
    
    Exit Function
Handler:
    ExistInCol = False

End Function
