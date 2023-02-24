VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPFrameDimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_guid As clsid
Private m_count As Long

Public Property Get Guid() As clsid
   LSet Guid = m_guid
End Property
Public Property Get GuidString() As String
   GuidString = GetGuidString(m_guid)
End Property

Public Property Get Name() As String
Dim sGuid As String
   sGuid = GetGuidString(m_guid)
   Select Case sGuid
   Case FrameDimensionPage
      Name = "Page"
   Case FrameDimensionResolution
      Name = "Resolution"
   Case FrameDimensionTime
      Name = "Time"
   Case Else
      Name = "Unknown Frame Dimension"
   End Select
End Property

Public Property Get FrameCount() As Long
   FrameCount = m_count
End Property

Friend Sub fInit(ByVal nativeImage As Long, dimensionId As clsid)
   LSet m_guid = dimensionId
   SetStatusHelper GdipImageGetFrameCount(nativeImage, _
      dimensionId, _
      m_count)
End Sub
