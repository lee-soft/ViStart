VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Handle As Long
Private m_hr As Long
Private m_pidl As SHFILEINFOW

Public IconPath As String
'Public IconFlags As Long

Public Function DrawIconEx(ByVal destinationHdc As Long, X As Long, Y As Long, cxIcon As Long, cyIcon As Long)
    win.DrawIconEx destinationHdc, X, Y, m_Handle, cxIcon, cyIcon, 0, 0, DI_NORMAL
End Function

Public Function DrawIcon(ByVal destinationHdc As Long, X As Long, Y As Long)
    'ImageList_Draw m_hr, m_Handle, destinationHdc, X, Y, ILD_TRANSPARENT
    
    win.DrawIconEx destinationHdc, X, Y, m_Handle, 16, 16, 0, 0, DI_NORMAL
End Function

Public Property Let Handle(newHandle As Long)
    FlushIcon
    m_Handle = newHandle
End Property

Public Property Get Handle() As Long
    Handle = m_Handle
End Property

Public Function FlushIcon()
    Class_Terminate
End Function

Private Sub Class_Terminate()
    If m_Handle <> 0 Then
        DestroyIcon ByVal m_Handle
    End If
End Sub
