VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : GDIPFont
'    Description: The Font class encapsulates the characteristics, such as family,
'    height, size, and style (or combination of styles), of a specific font.
'    A Font object is used when drawing strings.
'
'    Based on Microsoft's Font class
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534437(v=vs.85).aspx
'--------------------------------------------------------------------------------


Option Explicit

Private m_FontName   As String

Private m_fontFamily As GDIPFontFamily

Private m_ptrFont    As Long

Private m_Size       As Single

Private m_lastResult As Long

Private Sub Class_Initialize()
    Set m_fontFamily = New GDIPFontFamily
    '
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       Constructor
' Description:       Initializes a GDI+Font with 1 simple call with a font family
'                    ** Adhears to MSDN **
'
' Parameters :       r_FontFamily (GDIPFontFamily)
'                    r_Size (Single)
'                    fontStyle (fontStyle = FontStyleItalic)
Public Sub Constructor(ByRef r_FontFamily As GDIPFontFamily, _
                       ByVal r_Size As Single, _
                       Optional ByVal fontStyle As fontStyle = FontStyleItalic)
                       
    Set m_fontFamily = r_FontFamily
    m_Size = r_Size
    
    SetStatus GdipCreateFont(m_fontFamily.Value, m_Size, fontStyle, UnitPixel, m_ptrFont)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       Constructor2 - Kept for compatability reasons, discontinue
'                    use in favor of the official constructor

' Description:       Initializes this object based on a FontFamily object,
'                    a size, a font style, and a unit of measurement
' Parameters :       theFontFamily (Long)
'                    r_Size (Single)
'                    fontStyle (fontStyle = FontStyleItalic)
'                    unit (GpUnit = UnitPixel)
'--------------------------------------------------------------------------------
Public Sub Depreciated_Constructor2(ByRef theFontFamily As Long, _
                        ByVal r_Size As Single, _
                        Optional ByVal fontStyle As fontStyle = FontStyleItalic, _
                        Optional ByVal unit As GpUnit = UnitPixel)
    Dispose
    
    SetStatus GdipCreateFont(theFontFamily, r_Size, fontStyle, unit, m_ptrFont)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       Constructor - Kept for compatability reasons, discontinue
'                    use in favor of the official constructor

' Description:       Initializes a GDI+Font with 1 simple call with a font name
'                    IE: Constructor "Tahoma", 14. Saves creating a fontfamily
'                    TODO: Discontinue this method in favour of Constructor2
'
' Parameters :       r_FontName (String)
'                    r_Size (Single)
'                    fontStyle (fontStyle = FontStyleItalic)
'--------------------------------------------------------------------------------
Public Sub Depreciated_Constructor(ByRef r_FontName As String, _
                       ByRef r_Size As Single, _
                       Optional ByVal fontStyle As fontStyle = FontStyleItalic)
    Dispose
    
    m_FontName = r_FontName
    m_Size = r_Size
    
    m_fontFamily.Constructor r_FontName
    SetStatus GdipCreateFont(m_fontFamily.Value, m_Size, fontStyle, UnitPixel, m_ptrFont)
End Sub

Public Property Get Size() As Single
    Size = m_Size
End Property

Public Property Let Size(ByRef r_Size As Single)
    m_Size = r_Size
End Property

Public Property Get fontFamily()
    fontFamily = m_fontFamily.Value
End Property

Public Property Get Value()

    If m_ptrFont = 0 Then
        Exit Property

    End If
    
    Value = m_ptrFont
End Property

Public Sub Dispose()
    If Not (m_ptrFont = 0) Then
        GdipDeleteFont m_ptrFont
        m_ptrFont = 0
    End If

End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

Private Function SetStatus(ByVal status As GpStatus) As GpStatus
    m_lastResult = SetStatusHelper(status)
End Function
