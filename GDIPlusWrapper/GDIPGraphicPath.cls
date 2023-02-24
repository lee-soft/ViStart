VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPGraphicPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : GDIPGraphicPath

'    Description: A GraphicsPath object stores a sequence of lines, curves, and
'    shapes. You can draw the entire sequence by calling the DrawPath method of
'    a Graphics object. You can partition the sequence of lines, curves, and
'    shapes into figures, and with the help of a GraphicsPathIterator object,
'    you can draw selected figures. You can also place markers in the sequence,
'    so that you can draw selected portions of the path.
'
'    Based on Microsoft GraphicsPath class  :
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534456(v=vs.85).aspx
'--------------------------------------------------------------------------------
Option Explicit

Private m_PathPtr    As Long

Private m_lastResult As GpStatus

Friend Property Get Value()
    Value = m_PathPtr
End Property


'--------------------------------------------------------------------------------
' Procedure  :       AddLine
' Description:       The AddLine method adds a line to the current figure of
'                    this path.
' Parameters :       X1 (Single)
'                    Y1 (Single)
'                    X2 (Single)
'                    Y2 (Single)
'--------------------------------------------------------------------------------
Public Sub AddLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
    
    SetStatus GdipAddPathLine(m_PathPtr, X1, Y1, X2, Y2)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       AddRectangle
' Description:       The AddRectangle method adds a rectangle to
'                    this path.
' Parameters :       X (Single)
'                    y (Single)
'                    Width (Single)
'                    Height (Single)
'--------------------------------------------------------------------------------
Public Sub AddRectangle(X As Single, Y As Single, Width As Single, Height As Single)
    
    SetStatus GdipAddPathRectangleI(m_PathPtr, X, Y, Width, Height)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       AddArc
' Description:       The AddArc method adds an elliptical arc to the current
'                    figure of this path.
' Parameters :       X (Single)
'                    y (Single)
'                    Width (Single)
'                    Height (Single)
'                    startAngle (Single)
'                    sweepAngle (Single)
'--------------------------------------------------------------------------------
Public Sub AddArc(X As Single, _
                  Y As Single, _
                  Width As Single, _
                  Height As Single, _
                  startAngle As Single, _
                  sweepAngle As Single)
    
    SetStatus GdipAddPathArc(m_PathPtr, X, Y, Width, Height, startAngle, sweepAngle)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       AddEllipse
' Description:       The AddEllipse method adds an ellipse to this path.
' Parameters :       R (RECTF)
'--------------------------------------------------------------------------------
Public Sub AddEllipse(r As RECTF)
    SetStatus GdipAddPathEllipse(m_PathPtr, r.Left, r.Top, r.Width, r.Height)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       AddString
' Description:       The AddString method adds the outline of a string to this path.
' Parameters :       s (String)
'                    family (GDIPFontFamily)
'                    style (Long)
'                    emSize (Single)
'                    origin (RECTF)
'                    format (StringFormatFlags)
'--------------------------------------------------------------------------------
Public Sub AddString(s As String, _
                     family As GDIPFontFamily, _
                     Style As Long, _
                     emSize As Single, _
                     origin As RECTF, _
                     format As StringFormatFlags)
    
    SetStatus GdipAddPathString(m_PathPtr, s, -1, family.Value, Style, emSize, origin, format)
End Sub

Public Sub Constructor(Optional BrushMode As FillMode = FillModeWinding)
    Dispose
    SetStatus GdipCreatePath(BrushMode, m_PathPtr)
End Sub

Public Sub Dispose()

    If Not (m_PathPtr = 0) Then
        GdipDeletePath m_PathPtr
        m_PathPtr = 0
    End If

End Sub

Private Sub SetStatus(ByVal status As GpStatus)
    m_lastResult = SetStatusHelper(status)
End Sub

Private Sub Class_Initialize()
    Constructor
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub
