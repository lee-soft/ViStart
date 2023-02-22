VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : GDIPPen
'    Description: A Pen object is a Windows GDI+ object used to draw lines
'                 and curves.

'    Based on Microsoft's GDI+ Pen class :
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534485(v=vs.85).aspx
'--------------------------------------------------------------------------------


Option Explicit

Private m_hPen              As Long

Private m_Size              As Long 'local copy

Private m_Alpha             As Byte

Private WithEvents m_Colour As Colour 'local copy
Attribute m_Colour.VB_VarHelpID = -1

Public Property Get Value() As Long
    Value = m_hPen
End Property

Public Property Let Colour(ByRef r_Colour As Colour)
    Set m_Colour = r_Colour
End Property

Public Property Get Colour() As Colour
    Set Colour = m_Colour
End Property

Public Property Let Size(ByRef r_Size As Long)
    m_Size = r_Size
End Property

Public Property Get Size() As Long
    Size = m_Size
End Property

Public Property Get Alpha() As Byte
    Alpha = m_Alpha
End Property

Public Property Let Alpha(ByRef r_Alpha As Byte)
    m_Alpha = r_Alpha
End Property

'--------------------------------------------------------------------------------
' Procedure  :       Class_Initialize
' Description:       Creates a default pen with full visability
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Colour = New Colour
    m_Alpha = 255
    
    Call GdipCreatePen1(GetRGB_VB2GDIP(m_Colour.Value, m_Alpha), m_Size, UnitPixel, m_hPen)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       Constructor
' Description:       Call this straight after instancing the GDIPen object to
'                    set it up
'
' Parameters :       r_Colour (Colour)
'                    r_Size (Long)
'                    r_Alpha (Byte)
'--------------------------------------------------------------------------------
Public Sub Constructor(ByRef r_Colour As Colour, _
                       ByRef r_Size As Long, _
                       ByRef r_Alpha As Byte)
    Dispose

    Set m_Colour = r_Colour
    m_Size = r_Size
    m_Alpha = r_Alpha
    
    Call GdipCreatePen1(GetRGB_VB2GDIP(m_Colour.Value, m_Alpha), m_Size, UnitPixel, m_hPen)
End Sub

Public Sub Dispose()

    If Not (m_hPen = 0) Then
        GdipDeletePen m_hPen
        m_hPen = 0
    End If

End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       m_Colour_onChanged
' Description:       Automatically recreates the Pen when the colour is changed
' Parameters :
'--------------------------------------------------------------------------------
Private Sub m_Colour_onChanged()

    GdipDeletePen m_hPen
    Call GdipCreatePen1(GetRGB_VB2GDIP(m_Colour.Value), m_Size, UnitPixel, m_hPen)
End Sub
