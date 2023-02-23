VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : GDIPBrush
'    Description: Defines a brush of a single color. Brushes are used to
'    fill graphics shapes, such as rectangles, ellipses, pies, polygons, and
'    paths.
'
'    Based on Microsoft's SolidBrush Class   :
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534439(v=vs.85).aspx
'--------------------------------------------------------------------------------


Option Explicit

Private m_hBrush            As Long

Private WithEvents m_Colour As Colour 'local copy
Attribute m_Colour.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set m_Colour = New Colour 'perhaps the only time the alpha value of colour is used
    Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_Colour.Value, m_Colour.Alpha), m_hBrush)
End Sub

Public Property Get Value()
    Value = m_hBrush
End Property

Public Property Get Colour() As Colour
    Set Colour = m_Colour
End Property

Public Property Let Colour(ByRef r_Colour As Colour)
    Set m_Colour = r_Colour
    m_Colour_onChanged
End Property

Public Sub Dispose()

    If (m_hBrush = 0) Then
        GdipDeleteBrush m_hBrush
        m_hBrush = 0
    End If

End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

Private Sub m_Colour_onChanged()
    Dispose
    Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_Colour.Value, m_Colour.Alpha), m_hBrush)
End Sub
