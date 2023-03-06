VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

Private m_filePath As String
Private m_fileNumber As Integer

Public Sub OpenStream(ByVal filePath As String, Optional ByVal appendMode As Boolean = False)
    m_filePath = filePath
    m_fileNumber = FreeFile
    
    If appendMode Then
        If Dir(filePath) <> "" Then
            Open m_filePath For Append As #m_fileNumber
        Else
            Open m_filePath For Output As #m_fileNumber
        End If
    Else
        Open m_filePath For Output As #m_fileNumber
    End If
End Sub

Public Sub WriteText(ByVal text As String)
    Write #m_fileNumber, text;
End Sub

Public Sub WriteTextLine(ByVal text As String)
    Print #m_fileNumber, text
End Sub

Public Sub CloseStream()
    Close #m_fileNumber
End Sub

Public Sub Flush()
    Dim currentPosition As Long
    currentPosition = Seek(m_fileNumber)
    FlushFileBuffers m_fileNumber
    Seek m_fileNumber, currentPosition
End Sub
