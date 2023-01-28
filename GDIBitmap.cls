VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Bitmap As Long
Private m_bitmapInfo As BITMAP

Public Property Get Height()
    Height = m_bitmapInfo.bmHeight
End Property

Public Property Get Width()
    Width = m_bitmapInfo.bmWidth
End Property

Private Function GetBitmapInfo()
    If m_Bitmap = 0 Then Exit Function

    ' call the API
    GetObjectAPI m_Bitmap, Len(m_bitmapInfo), m_bitmapInfo
End Function

Public Function CreateCompatibleBitmapFromHDC(hdc As Long, Width As Long, Height As Long)
    m_Bitmap = CreateCompatibleBitmap(hdc, Width, Height)
    GetBitmapInfo
    
    If m_Bitmap = 0 Then
        MsgBox ":("
    End If
End Function

Public Function LoadImageFromFile(strSource As String)

    m_Bitmap = LoadImage(App.hInstance, strSource, IMAGE_BITMAP, 0, 0, _
                            LR_COLOR Or LR_LOADFROMFILE)
    GetBitmapInfo
End Function

Public Property Get hBitmap() As Long
    hBitmap = m_Bitmap
End Property

Public Property Let hBitmap(new_hBitmap As Long)
    Class_Terminate
    m_Bitmap = new_hBitmap
    
End Property

Private Sub Class_Terminate()
    If m_Bitmap <> 0 Then
        DeleteObject m_Bitmap
    End If
End Sub
