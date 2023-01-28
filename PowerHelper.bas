Attribute VB_Name = "PowerHelper"
Option Explicit

'An attempt to tidy up API
'One day i will place all user32 api functions here
Public Const EWX_LOGOFF As Long = 0
Public Const EWX_POWEROFF As Long = &H8
Public Const EWX_REBOOT As Long = 2
Public Const EWX_FORCEIFHUNG As Long = &H10

Public Declare Function SetSuspendState Lib "powrprof.dll" (ByVal Hibernate As Long, ByVal ForceCritical As Long, ByVal DisableWakeEvent As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

' ===========================================================
' NT Only
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(0 To 0) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" _
    (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
     (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
     NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
     PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
     Alias "LookupPrivilegeValueA" _
    (ByVal lpSystemName As String, ByVal lpName As String, _
    lpLuid As LUID) As Long
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const SE_PRIVILEGE_ENABLED = &H2

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = (&H2)
Private Const TOKEN_IMPERSONATE = (&H4)
Private Const TOKEN_QUERY = (&H8)
Private Const TOKEN_QUERY_SOURCE = (&H10)
Private Const TOKEN_ADJUST_PRIVILEGES = (&H20)
Private Const TOKEN_ADJUST_GROUPS = (&H40)
Private Const TOKEN_ADJUST_DEFAULT = (&H80)
Private Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
                        TOKEN_ASSIGN_PRIMARY Or _
                        TOKEN_DUPLICATE Or _
                        TOKEN_IMPERSONATE Or _
                        TOKEN_QUERY Or _
                        TOKEN_QUERY_SOURCE Or _
                        TOKEN_ADJUST_PRIVILEGES Or _
                        TOKEN_ADJUST_GROUPS Or _
                        TOKEN_ADJUST_DEFAULT)
Private Const TOKEN_READ = (STANDARD_RIGHTS_READ Or _
                        TOKEN_QUERY)
Private Const TOKEN_WRITE = (STANDARD_RIGHTS_WRITE Or _
                        TOKEN_ADJUST_PRIVILEGES Or _
                        TOKEN_ADJUST_GROUPS Or _
                        TOKEN_ADJUST_DEFAULT)
Private Const TOKEN_EXECUTE = (STANDARD_RIGHTS_EXECUTE)

Private Const TokenDefaultDacl = 6
Private Const TokenGroups = 2
Private Const TokenImpersonationLevel = 9
Private Const TokenOwner = 4
Private Const TokenPrimaryGroup = 5
Private Const TokenPrivileges = 3
Private Const TokenSource = 7
Private Const TokenStatistics = 10
Private Const TokenType = 8
Private Const TokenUser = 1

Public Function NTEnableShutDown(ByRef sMsg As String) As Boolean

Dim tLUID As LUID
Dim hProcess As Long
Dim hToken As Long
Dim tTP As TOKEN_PRIVILEGES, tTPOld As TOKEN_PRIVILEGES
Dim lTpOld As Long
Dim lR As Long

    ' Under NT we must enable the SE_SHUTDOWN_NAME privilege in the
    ' process we're trying to shutdown from, otherwise a call to
    ' try to shutdown has no effect!

    ' Find the LUID of the Shutdown privilege token:
    lR = LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, tLUID)
    
    ' If we get it:
    If (lR <> 0) Then
                
       ' Get the current process handle:
       hProcess = GetCurrentProcess()
       If (hProcess <> 0) Then
           ' Open the token for adjusting and querying
           ' (if we can - user may not have rights):
           lR = OpenProcessToken(hProcess, _
                   TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
           If (lR <> 0) Then
                       
               ' Ok we can now adjust the shutdown priviledges:
               With tTP
                   .PrivilegeCount = 1
                   With .Privileges(0)
                      .Attributes = SE_PRIVILEGE_ENABLED
                      .pLuid.HighPart = tLUID.HighPart
                      .pLuid.LowPart = tLUID.LowPart
                   End With
               End With
            
               ' Now allow this process to shutdown the system:
               lR = AdjustTokenPrivileges(hToken, 0, tTP, Len(tTP), tTPOld, lTpOld)
            
               If (lR <> 0) Then
                  NTEnableShutDown = True
               Else
                  MsgBox "Cant enable shutdown, You do not have sufficient access " & _
                    "to shutdown windows.", vbCritical
               End If
            
               ' Remember to close the handle when finished with it:
               CloseHandle hToken
           Else
               MsgBox "Cant enable shutdown, You do not have sufficient access " & _
                 " to shutdown windows.", vbCritical
           End If
       Else
           MsgBox "Cant enable shutdown, Cant find the process.", vbCritical
       End If
    Else
       MsgBox "Cant enable shutdown, unknown error.", vbCritical
    End If


End Function
