VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPFontFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : GDIPFontFamily

'    Description: This FontFamily class encapsulates a set of fonts that make up
'    a font family. A font family is a group of fonts that have the same
'    typeface but different styles.
'
'    Based on Microsoft's GDI+ FontFamily class  :
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534439(v=vs.85).aspx
'--------------------------------------------------------------------------------


Option Explicit

Private m_FontFamilyPtr As Long

Private m_lastResult    As GpStatus

Friend Property Get Value()
    Value = m_FontFamilyPtr
End Property

'--------------------------------------------------------------------------------
' Procedure  :       IsStyleAvailable
' Description:       The IsAvailable method determines whether the specified
'                    style is available for this font family.
' Parameters :       style (gdiplus.fontStyle)
'--------------------------------------------------------------------------------
Public Function IsStyleAvailable(ByVal Style As gdiplus.fontStyle) As Boolean

    Dim lngStyleAvailable As Long

    SetStatus GdipIsStyleAvailable(m_FontFamilyPtr, Style, lngStyleAvailable)
    
    If m_lastResult = gdiplus.Ok Then
        IsStyleAvailable = lngStyleAvailable
    Else
        IsStyleAvailable = False
    End If

End Function

'--------------------------------------------------------------------------------
' Procedure  :       Clone
' Description:       The Clone method creates a new FontFamily object based on
'                    this FontFamily object.
' Parameters :
'--------------------------------------------------------------------------------
Public Function Clone() As GDIPFontFamily

    Dim lngFamily As Long

    Dim clsClone  As GDIPFontFamily
    
    SetStatus GdipCloneFontFamily(m_FontFamilyPtr, lngFamily)

    If m_lastResult <> Ok Then Set clsClone = Nothing: Exit Function
    Set clsClone = New GDIPFontFamily
    clsClone.fInitialize (lngFamily)
    Set Clone = clsClone
End Function

Public Sub Dispose()

    If Not (m_FontFamilyPtr = 0) Then
        GdipDeleteFontFamily m_FontFamilyPtr
        m_FontFamilyPtr = 0
    End If

End Sub

Friend Sub fInitialize(ByVal lFamily As Long)
    Dispose
    m_FontFamilyPtr = lFamily
End Sub

Public Sub Constructor(Optional FontName As String = "Arial")
    Dispose
    SetStatus GdipCreateFontFamilyFromName(FontName, 0, m_FontFamilyPtr)
End Sub

Public Sub Constructor2(ByVal Name As String, ByVal NativeFontCollection As Long)
    Dispose
    SetStatus GdipCreateFontFamilyFromName(Name, NativeFontCollection, m_FontFamilyPtr)
End Sub

Private Sub Class_Initialize()
    Constructor
End Sub

Private Sub SetStatus(ByVal status As GpStatus)
    m_lastResult = SetStatusHelper(status)
End Sub
