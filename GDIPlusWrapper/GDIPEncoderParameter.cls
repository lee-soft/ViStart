VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPEncoderParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_guid As clsid
Private m_type As EncoderParameterValueType
Private m_value() As Variant
Private m_iValueCount As Long
Private m_allowableValues() As Variant
Private m_iAllowableValueCount As Long

Public Property Get Value(ByVal index As Long) As Variant
   Value = m_value(index)
End Property
Public Property Let Value(ByVal index As Long, ByVal theValue As Variant)
   m_value(index) = theValue
End Property
Public Property Get ValueCount() As Long
   ValueCount = m_iValueCount
End Property

Public Property Let ValueCount(ByVal count As Long)
   m_iValueCount = count
   If (count > 0) Then
      ReDim Preserve m_value(1 To m_iValueCount) As Variant
   Else
      Erase m_value
   End If
End Property

Public Property Get AllowableValueCount() As Long
   AllowableValueCount = m_iAllowableValueCount
End Property
Public Property Get AllowableValue(ByVal index As Long)
   AllowableValue = m_allowableValues(index)
End Property

Public Property Get Guid() As clsid
   LSet Guid = m_guid
End Property

Public Property Get GuidString() As String
   GuidString = GetGuidString(m_guid)
End Property

Public Property Get Name() As String
Dim sGuid As String

   sGuid = GetGuidString(m_guid)
   
   If (StrComp(sGuid, EncoderCompression, vbTextCompare) = 0) Then
      Name = "Compression"
   ElseIf (StrComp(sGuid, EncoderColorDepth, vbTextCompare) = 0) Then
      Name = "ColorDepth"
   ElseIf (StrComp(sGuid, EncoderScanMethod, vbTextCompare) = 0) Then
      Name = "ScanMethod"
   ElseIf (StrComp(sGuid, EncoderVersion, vbTextCompare) = 0) Then
      Name = "Version"
   ElseIf (StrComp(sGuid, EncoderRenderMethod, vbTextCompare) = 0) Then
      Name = "RenderMethod"
   ElseIf (StrComp(sGuid, EncoderQuality, vbTextCompare) = 0) Then
      Name = "Quality"
   ElseIf (StrComp(sGuid, EncoderTransformation, vbTextCompare) = 0) Then
      Name = "Transformation"
   ElseIf (StrComp(sGuid, EncoderLuminanceTable, vbTextCompare) = 0) Then
      Name = "LuminanceTable"
   ElseIf (StrComp(sGuid, EncoderChrominanceTable, vbTextCompare) = 0) Then
      Name = "ChrominanceTable"
   ElseIf (StrComp(sGuid, EncoderSaveFlag, vbTextCompare) = 0) Then
      Name = "SaveFlag"
   Else
      Name = "Unknown category"
   End If
   
End Property


Public Property Get ParamType() As EncoderParameterValueType
   ParamType = m_type
End Property

Friend Function fInit( _
      iguid As clsid, _
      iNumValues As EncoderParameterValueType, _
      itype As Long, _
      lPtr As Long _
   )
   LSet m_guid = iguid
   m_iAllowableValueCount = iNumValues
   m_type = itype
   If (m_iAllowableValueCount > 0) Then
      Dim i As Long
      ' Get the allowable values (currently only Long and LongRange do anything):
      If (m_type = EncoderParameterValueTypeLong) Then
         ReDim l(0 To m_iAllowableValueCount - 1) As Long
         ReDim m_allowableValues(1 To m_iAllowableValueCount) As Variant
         RtlMoveMemory ByVal VarPtr(l(0)), ByVal lPtr, m_iAllowableValueCount * 4
         For i = 1 To m_iAllowableValueCount
            m_allowableValues(i) = l(i - 1)
         Next i
      ElseIf (m_type = EncoderParameterValueTypeLongRange) Then
         ReDim m_allowableValues(1 To 1) As Variant
         ReDim l(1 To 2) As Long
         RtlMoveMemory ByVal VarPtr(l(1)), ByVal lPtr, 8
         m_allowableValues(1) = l
      End If
   End If
End Function

