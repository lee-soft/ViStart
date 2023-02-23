VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPPathGradientBrush"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : GDIPPathGradientBrush
'    Description: Encapsulates a Brush object that fills the interior of a
'                 GraphicsPath object with a gradient.
'
'    Based on Microsoft's GDI+ PathGradiantBrush class :
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534483(v=vs.85).aspx
'--------------------------------------------------------------------------------


Private m_PathPtr      As Long

Private m_lastResult   As Long

Private m_centerColour As Colour

Friend Property Get Value()
    Value = m_PathPtr
End Property


'--------------------------------------------------------------------------------
' Procedure  :       SurroundColours
' Description:       Sets an array of colors that correspond to the points in
'                    the path this PathGradientBrush fills.
'                    TODO: Turn into a property
' Parameters :       theColours() (Colour)
'--------------------------------------------------------------------------------
Public Function SurroundColours(theColours() As Colour)

    Dim colourArr() As Long

    Dim thisColour  As Colour

    Dim colourIndex As Long

    ReDim colourArr(UBound(theColours))

    For colourIndex = LBound(theColours) To UBound(theColours)
        Set thisColour = theColours(colourIndex)
        colourArr(colourIndex) = GetRGB_VB2GDIP(thisColour.Value, thisColour.Alpha)
    Next

    GdipSetPathGradientSurroundColorsWithCount m_PathPtr, colourArr(0), UBound(colourArr) + 1
End Function


'--------------------------------------------------------------------------------
' Procedure  :       CenterColour
' Description:       Gets the color at the center of the path gradient.
' Parameters :
'--------------------------------------------------------------------------------
Public Property Get CenterColour() As Colour
    Set Colour = m_centerColour
End Property

'--------------------------------------------------------------------------------
' Procedure  :       CenterColour
' Description:       Lets the color at the center of the path gradient.
' Parameters :       r_Colour (Colour)
'--------------------------------------------------------------------------------
Public Property Let CenterColour(ByRef r_Colour As Colour)
    Set m_centerColour = r_Colour
    GdipSetPathGradientCenterColor m_PathPtr, GetRGB_VB2GDIP(r_Colour.Value, r_Colour.Alpha)
End Property

Public Sub Constructor(points() As POINTF)
    Dispose
    SetStatus GdipCreatePathGradient(points(0), UBound(points) + 1, WrapModeTileFlipXY, m_PathPtr)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       FromPath
' Description:       Creates a GradiantBrush from a GraphicPath object
' Parameters :       sourcePath (GDIPGraphicPath)
'--------------------------------------------------------------------------------
Public Sub FromPath(ByRef sourcePath As GDIPGraphicPath)
    Dispose
    SetStatus GdipCreatePathGradientFromPath(sourcePath.Value, m_PathPtr)
End Sub

Public Sub Dispose()

    If Not (m_PathPtr = 0) Then
        GdipDeletePath m_PathPtr
        m_PathPtr = 0
    End If

End Sub

Private Sub Class_Initialize()
    Set m_centerColour = New Colour
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

Private Sub SetStatus(ByVal status As GpStatus)
    m_lastResult = SetStatusHelper(status)
End Sub
