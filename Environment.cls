VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long

Private m_version As VersionInfo

Private m_logger As SeverityLogger

Property Get Logger() As SeverityLogger
    Set Logger = m_logger
End Property

Public Property Get OSVersion() As VersionInfo
    Set OSVersion = m_version
End Property

Public Function GetEnvironmentVariable(variableName As String) As String
    GetEnvironmentVariable = Environ$(variableName)
End Function

Private Sub Class_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)
    Set m_version = New VersionInfo

    Dim kernalPath As String: kernalPath = Environ("windir") & "\System32\ntoskrnl.exe"
    Dim kernalFileInfo As FileVersionInfo: Set kernalFileInfo = FileVersionInfoHelper.GetVersionInfo(kernalPath)

    m_version.Major = kernalFileInfo.ProductMajorPart
    m_version.Minor = kernalFileInfo.ProductMinorPart
    m_version.Build = kernalFileInfo.ProductBuildPart
    
    If kernalFileInfo.ProductMajorPart = 10 And kernalFileInfo.ProductBuildPart >= 22000 Then
        ' Windows 11
        m_version.Minor = 5

    ElseIf kernalFileInfo.ProductMajorPart = 10 Then
        ' Windows 10
        m_version.Minor = 4

    ElseIf kernalFileInfo.ProductMajorPart = 12 Then
        ' Windows >11
        m_version.Minor = 6
    End If

End Sub

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
