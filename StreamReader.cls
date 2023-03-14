VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_filePath As String
Private m_fileNumber As Integer

Public Sub OpenStream(ByVal filePath As String)
    m_filePath = filePath
    m_fileNumber = FreeFile
    
    Open m_filePath For Input As #m_fileNumber
End Sub

Public Function ReadToEnd()
Dim stringData As String
    stringData = input(LOF(m_fileNumber), #m_fileNumber)
    
    ReadToEnd = stringData
End Function


Public Sub CloseStream()
    Close #m_fileNumber
End Sub

Private Sub Class_Terminate()
    CloseStream
End Sub
