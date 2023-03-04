VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : LogManager
'    Description: Creates and manages instances of Logger objects.
'                 Call LogManager.GetCurrentClassLogger(Me) from any class/form
'                 To enable severity logging on the object
'
'    Author     : Lee Chantrey
'--------------------------------------------------------------------------------

Option Explicit

Private m_defaultLogger As SeverityLogger
Private m_loggers As Collection

Public Enum LogLevel
    TraceLevel = 1
    InfoLevel = 2
    WarnLevel = 3
    ErrorLevel = 4
    FatalLevel = 5
End Enum

Public Enum TargetType
    LogFile
    LogDebug
    LogAll
End Enum

Private m_logLevel As LogLevel
Private m_targetType As TargetType
Private m_logFile As StreamWriter
Private m_autoFlush As Boolean

Public Property Let AutoFlush(ByVal newAutoFlush As Boolean)
    m_autoFlush = newAutoFlush
End Property

Public Property Let Target(newTargetType As TargetType)
    If newTargetType = LogFile Or newTargetType = LogAll Then
        Set m_logFile = New StreamWriter
        m_logFile.OpenStream App.Path & "\errors.log", True
    Else
        Set m_logFile = Nothing
    End If
    
    m_targetType = newTargetType
End Property

Public Property Let MinimumLevel(newLevel As LogLevel)
    m_logLevel = newLevel
End Property

Public Function InvokeLogger(ByVal level As LogLevel, ByVal Source As String, ByVal message As String, Optional arguments As Variant)
    If level < m_logLevel Then
        Exit Function
    End If
    
Dim sourceHeader As String

    If Not IsMissing(arguments) Then
        sourceHeader = Source & " " & Join(arguments, ",")
    Else
        sourceHeader = Source
    End If
    
    If m_targetType = LogDebug Or m_targetType = LogAll Then
        Debug.Print Now(); " "; level; " "; message; " "; sourceHeader
    End If
    
    If m_targetType = LogFile Or m_targetType = LogAll Then
        If Not m_logFile Is Nothing Then
            m_logFile.WriteTextLine Now() & " " & level & " " & message & " " & sourceHeader
            If m_autoFlush Then m_logFile.Flush
        End If
    End If
End Function

Public Function GetLogger(ByVal loggerKey As String) As SeverityLogger

Dim targetLogger As SeverityLogger

    If Not IsLogged(loggerKey) Then
        Set targetLogger = New SeverityLogger
        targetLogger.Initialize Me, loggerKey
        m_loggers.Add targetLogger, loggerKey
    End If
    
    Set targetLogger = m_loggers(loggerKey)
    Set GetLogger = targetLogger

End Function

Public Function GetCurrentClassLogger(ByRef sourceObject As Object) As SeverityLogger
    Set GetCurrentClassLogger = m_defaultLogger

Dim objectKey As String
    objectKey = TypeName(sourceObject)

    If objectKey = vbNullString Then
        Exit Function
    End If
    
    Set GetCurrentClassLogger = GetLogger(objectKey)
End Function

Private Sub Class_Initialize()
    Set m_defaultLogger = New SeverityLogger
    Set m_loggers = New Collection
    
    m_logLevel = TraceLevel
    Target = LogAll
    AutoFlush = True
End Sub

Private Function IsLogged(loggerKey As String) As Boolean

    On Error GoTo Handler
    IsLogged = Not (IsEmpty(m_loggers(loggerKey)))
    
    Exit Function
Handler:
    IsLogged = False

End Function
