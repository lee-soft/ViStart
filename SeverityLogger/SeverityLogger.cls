VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeverityLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : SeverityLogger
'    Description: Provides NLog style severity logging logger.Info("Hello!")
'
'    Author     : Lee Chantrey
'--------------------------------------------------------------------------------


Option Explicit

Private m_factory As LogManager
Private m_sourceModule As String

Public Property Get SourceModule() As String
    SourceModule = m_sourceModule
End Property

Public Property Get Factory() As LogManager
    Set Factory = m_factory
End Property

Public Function Initialize(ByRef logFactory As LogManager, ByVal newSourceModule As String)
    Set m_factory = logFactory
    m_sourceModule = newSourceModule
End Function

' Only when I would be "tracing" the code and trying to find one part of a
' function specifically.
Public Function Trace(ByVal traceMessage As String, ByVal sourceProcedure As String, ParamArray arguments() As Variant)
    
    Dim stringArray() As String
    Dim argumentIndex As Long
    
    If UBound(arguments) > -1 Then
        ReDim stringArray(0 To UBound(arguments)) ' Resize the string array to match the variant array size
        
        For argumentIndex = 0 To UBound(arguments)
            stringArray(argumentIndex) = CStr(arguments(argumentIndex)) ' Convert the variant to a string and add it to the string array
        Next argumentIndex
    End If
    
    m_factory.InvokeLogger TraceLevel, m_sourceModule & "::" & sourceProcedure, traceMessage, stringArray
End Function

' Generally useful information to log (service start/stop, configuration
' assumptions, etc). Info I want to always have available but usually don't
' care about under normal circumstances. This is my out-of-the-box config level.
Public Function Info(infoMessage As String)
    m_factory.InvokeLogger InfoLevel, m_sourceModule, infoMessage
End Function

' Anything that can potentially cause application oddities, but for which I am
' automatically recovering. (Such as switching from a primary to backup server,
' retrying an operation, missing secondary data, etc.)
Public Function Warn(warningMessage As String, ByVal sourceProcedure As String, ParamArray arguments() As Variant)
Dim stringArray() As String
    Dim argumentIndex As Long
    
    If UBound(arguments) > -1 Then
        ReDim stringArray(0 To UBound(arguments)) ' Resize the string array to match the variant array size
        
        For argumentIndex = 0 To UBound(arguments)
            stringArray(argumentIndex) = CStr(arguments(argumentIndex)) ' Convert the variant to a string and add it to the string array
        Next argumentIndex
    End If
 
    m_factory.InvokeLogger WarnLevel, m_sourceModule & "::" & sourceProcedure, warningMessage, stringArray
End Function

' Any error which is fatal to the operation, but not the service or application
' (can't open a required file, missing data, etc.). These errors will force user
' (administrator, or direct user) intervention. These are usually reserved
' (in my apps) for incorrect connection strings, missing services, etc.
Public Function Error(ByVal errorDescription As String, ByVal sourceProcedure As String, ParamArray arguments() As Variant)

    Dim stringArray() As String
    Dim argumentIndex As Long
    
    If UBound(arguments) > -1 Then
        ReDim stringArray(0 To UBound(arguments)) ' Resize the string array to match the variant array size
        
        For argumentIndex = 0 To UBound(arguments)
            stringArray(argumentIndex) = CStr(arguments(argumentIndex)) ' Convert the variant to a string and add it to the string array
        Next argumentIndex
    End If
    
    m_factory.InvokeLogger ErrorLevel, m_sourceModule & "::" & sourceProcedure, errorDescription, stringArray
End Function

' Any error that is forcing a shutdown of the service or application to prevent
' data loss (or further data loss). I reserve these only for the most heinous
' errors and situations where there is guaranteed to have been data corruption
' or loss.
Public Function Fatal(errorDescription As String, sourceProcedure As String, ParamArray arguments() As Variant)
    Dim stringArray() As String
    Dim argumentIndex As Long

    If UBound(arguments) > -1 Then
        ReDim stringArray(0 To UBound(arguments)) ' Resize the string array to match the variant array size
        
        For argumentIndex = 0 To UBound(arguments)
            stringArray(argumentIndex) = CStr(arguments(argumentIndex)) ' Convert the variant to a string and add it to the string array
        Next argumentIndex
    End If
    
    m_factory.InvokeLogger FatalLevel, m_sourceModule & "::" & sourceProcedure, errorDescription, stringArray
End Function
