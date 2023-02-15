VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' cMemDC - flicker free drawing
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private m_hDC As Long
Private m_hBmp As Long
Private m_hBmpOld As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_rect As RECT

Public Function Clear()
    FillRect m_hDC, m_rect, GetStockObject(WHITE_BRUSH)
End Function

Public Property Get Width() As Long
   Width = m_lWidth
End Property
Public Property Let Width(ByVal Value As Long)
Dim lJunk As Long
   If (Value > m_lWidth) Then
      m_lWidth = Value
      pCreate m_lWidth, m_lHeight
      SetBkMode m_hDC, 1
   End If
End Property
Public Property Get Height() As Long
   Height = m_lHeight
End Property
Public Property Let Height(ByVal Value As Long)
Dim lJunk As Long
   If (Value > m_lHeight) Then
      m_lHeight = Value
      pCreate m_lWidth, m_lHeight
      SetBkMode m_hDC, 1
   End If
End Property

Public Property Get hdc() As Long
   hdc = m_hDC
End Property
Public Sub Draw( _
      ByVal hdc As Long, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
   If WidthSrc <= 0 Then WidthSrc = m_lWidth
   If HeightSrc <= 0 Then HeightSrc = m_lHeight
   BitBlt hdc, xDst, yDst, WidthSrc, HeightSrc, m_hDC, xSrc, ySrc, vbSrcCopy
End Sub
Public Sub CreateFromPicture(sPic As IPicture)
Dim tB As BITMAP
Dim lhDCC As Long, lhDC As Long
Dim lhBmpOld As Long
   GetObjectAPI sPic.Handle, Len(tB), tB
   Width = tB.bmWidth
   Height = tB.bmHeight
   lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDC = CreateCompatibleDC(lhDCC)
   lhBmpOld = SelectObject(lhDC, sPic.Handle)
   BitBlt m_hDC, 0, 0, tB.bmWidth, tB.bmHeight, lhDC, 0, 0, vbSrcCopy
   SelectObject lhDC, lhBmpOld
   DeleteDC lhDC
   DeleteDC lhDCC
End Sub

Public Function CreateFromBitmap(sourceBitmap As Long) As Boolean

Dim hbmT As Long
Dim lngBitMapDC As Long

    lngBitMapDC = CreateCompatibleDC(0)
    hbmT = SelectObject(lngBitMapDC, sourceBitmap)
    
    m_hDC = lngBitMapDC
    CreateFromBitmap = True
End Function

Public Function CreateFromFile(strFile As String) As Boolean
    'Creates a bitmap hdc
    'Returns the bitmap hdc

    On Error GoTo Handler

Dim lngBitMapDC As Long
    lngBitMapDC = CreateCompatibleDC(0)
    
    SelectObject lngBitMapDC, LoadPicture(strFile)

    m_hDC = lngBitMapDC
    CreateFromFile = True
    
    Exit Function
Handler:
    DeleteDC m_hDC
    CreateFromFile = False
End Function

Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
Dim lhDCC As Long
   pDestroy
   
   lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
   If Not (lhDCC = 0) Then
      m_hDC = CreateCompatibleDC(lhDCC)
      If Not (m_hDC = 0) Then
         m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
         If Not (m_hBmp = 0) Then
            m_hBmpOld = SelectObject(m_hDC, m_hBmp)
            If Not (m_hBmpOld = 0) Then
               m_lWidth = Width
               m_lHeight = Height
               m_rect.Bottom = Height
               m_rect.Right = Width
               
               DeleteDC lhDCC
               Exit Sub
            End If
         End If
      End If
      DeleteDC lhDCC
      pDestroy
   End If
End Sub
Private Sub pDestroy()
   If Not m_hBmpOld = 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If Not m_hDC = 0 Then
      DeleteDC m_hDC
      m_hDC = 0
   End If
   m_lWidth = 0
   m_lHeight = 0
End Sub

Private Sub Class_Terminate()
   pDestroy
End Sub




