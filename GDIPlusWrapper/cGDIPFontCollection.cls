VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPFontCollection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================================
' cFontCollection.cls
' Copyright ? 2003 Genghis Khan(GuangJian Guo)
' Website: http://www.itkhan.com
'
' from GdiplusFontCollection.h:Microsoft Platform SDK February 2003
'
' The FontCollection class is an abstract base class. It contains methods for
' enumerating the font families in a collection of fonts. Objects built from this
' class include the InstalledFontCollection and PrivateFontCollection classes
' ==================================================================================
Option Explicit

Private m_lngFontCollection As Long

Private m_uLastStatus       As GpStatus

Friend Sub fInitialize(ByVal lCol As Long)
    m_lngFontCollection = lCol
End Sub

' gets the font families contained in this font collection.
Public Function GetFamilies(ByVal numSought As Long, _
                            ByRef gpFamilies() As GDIPFontFamily, _
                            ByRef numFound As Long) As GpStatus

    Dim i                  As Long

    Dim nativeFamilyList() As Long

    Dim lngClone           As Long
    
    If numSought <= 0 Then
        m_uLastStatus = InvalidParameter
        GetFamilies = m_uLastStatus

        Exit Function

    End If

    ReDim nativeFamilyList(numSought - 1)
    m_uLastStatus = GdipGetFontCollectionFamilyList(m_lngFontCollection, numSought, nativeFamilyList(0), numFound)

    If m_uLastStatus <> Ok Then Exit Function
    ReDim gpFamilies(0 To numSought - 1)

    For i = 0 To numSought - 1
        Set gpFamilies(i) = New GDIPFontFamily
        m_uLastStatus = GdipCloneFontFamily(nativeFamilyList(i), lngClone)

        If m_uLastStatus = Ok Then gpFamilies(i).fInitialize lngClone
    Next i

    Erase nativeFamilyList
    GetFamilies = m_uLastStatus
End Function

Public Function GetFamilyCount() As Long

    Dim lCount As Long

    m_uLastStatus = GdipGetFontCollectionFamilyCount(m_lngFontCollection, lCount)
    GetFamilyCount = lCount
End Function

' returns a value that indicates the nature of this CustomLineCap object's most recent method failure.
Public Function GetLastStatus() As GpStatus
    GetLastStatus = m_uLastStatus
End Function

Public Property Get NativeFontCollection() As Long
    NativeFontCollection = m_lngFontCollection
End Property

Private Sub Class_Initialize()
    m_lngFontCollection = 0
End Sub

Private Sub Class_Terminate()
    m_lngFontCollection = 0
End Sub
