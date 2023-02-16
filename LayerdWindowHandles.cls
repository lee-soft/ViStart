VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayerdWindowHandles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public theDC As Long
Public oldBitmap As Long
Public mainBitmap As Long

Private winSize As SIZEL
Private srcPoint As POINTL

Public Function SetSize(newSize As SIZEL)
    winSize = newSize
End Function

Public Function GetSize() As SIZEL
    GetSize = winSize
End Function

Public Function SetPoint(newPoint As POINTL)
    srcPoint = newPoint
End Function

Public Function GetPoint() As POINTL
    GetPoint = srcPoint
End Function

Private Sub Class_Terminate()
    SelectObject theDC, oldBitmap
    
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    'DeleteDC theDC
End Sub
