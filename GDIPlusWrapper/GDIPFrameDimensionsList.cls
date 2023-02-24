VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPFrameDimensionsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function LocalGdipImageGetFrameDimensionsList Lib "gdiplus.dll" Alias "GdipImageGetFrameDimensionsList" ( _
   ByVal Image As Long, _
   ByVal lPtrBuff As Long, _
   ByVal count As Long) As GpStatus

Private m_count As Long
Private m_cFrameDimensions() As GDIPFrameDimension

Public Property Get FrameDimension(ByVal index As Long) As GDIPFrameDimension
   If (index > 0) And (index <= m_count) Then
      Set FrameDimension = m_cFrameDimensions(index)
   Else
      SetStatusHelper InvalidParameter
   End If
End Property

Friend Sub fInit( _
      ByVal nativeImage As Long _
   )
   SetStatusHelper GdipImageGetFrameDimensionsCount(nativeImage, m_count)
   If (m_count > 0) Then
      ReDim m_cFrameDimensions(1 To m_count) As GDIPFrameDimension
      Dim lPtrBuff As Long
      Dim p As clsid
      ReDim b(0 To Len(p) * m_count) As Byte
      lPtrBuff = VarPtr(b(0))
      
      SetStatusHelper _
         LocalGdipImageGetFrameDimensionsList( _
            nativeImage, _
            lPtrBuff, _
            m_count)
            
       Dim i As Long
       Dim lStart As Long
       For i = 1 To m_count
         RtlMoveMemory p, ByVal b(lStart), Len(p)
         Set m_cFrameDimensions(i) = New GDIPFrameDimension
         m_cFrameDimensions(i).fInit nativeImage, p
         lStart = lStart + Len(p)
       Next i
   End If
   
End Sub

Public Property Get count() As Long
   count = m_count
End Property

