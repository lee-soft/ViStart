VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_DC As Long

Public Property Get Handle() As Long
    Handle = m_DC
End Property

Private Sub Class_Initialize()
    m_DC = CreateCompatibleDC(0)
    
End Sub

Public Function SelectBitmap(srcBitmap As GDIBitmap)
    SelectObject m_DC, srcBitmap.hBitmap
End Function

Private Sub Class_Terminate()
    If m_DC <> 0 Then
        DeleteDC m_DC
    End If
End Sub
