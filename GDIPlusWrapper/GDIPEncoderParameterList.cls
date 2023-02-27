VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPEncoderParameterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LocalGdipGetEncoderParameterList Lib "gdiplus.dll" Alias "GdipGetEncoderParameterList" ( _
   ByVal Image As Long, _
   clsidEncoder As clsid, _
   ByVal Size As Long, _
   ByVal lPtrBuff As Long) As GpStatus

Private m_clsIdEncoder As clsid
Private m_count As Long
Private m_cParameters() As GDIPEncoderParameter
Private m_lSize As Long

Public Property Get count() As Long
   count = m_count
End Property

Public Property Get clsidEncoder() As clsid
   LSet clsidEncoder = m_clsIdEncoder
End Property

Public Property Get Parameter(ByVal index As Long) As GDIPEncoderParameter
   If (index > 0) And (index <= m_count) Then
      Set Parameter = m_cParameters(index)
   Else
      SetStatusHelper InvalidParameter
   End If
End Property

Public Property Get ParameterForGuid(ByVal sGuid As String) As GDIPEncoderParameter
Dim i As Long
   For i = 1 To m_count
      If (StrComp(m_cParameters(i).GuidString, sGuid) = 0) Then
         Set ParameterForGuid = m_cParameters(i)
         Exit For
      End If
   Next i
End Property

Private Function pEvaluateBufSize() As Long
Dim lSize As Long
Dim i As Long
Dim pP As EncoderParameter
   lSize = 0
   If (m_count > 0) Then
      lSize = lSize + 4 ' count
      For i = 1 To m_count
         If (m_cParameters(i).ValueCount > 0) Then
            lSize = lSize + Len(pP)  ' size of item
         End If
      Next i
   End If
   pEvaluateBufSize = lSize
End Function

Friend Function fGetEncoderParameterBufSize() As Long
   '
   fGetEncoderParameterBufSize = pEvaluateBufSize()
   '
End Function

Friend Sub fInit(cBitmap As GDIPImage, clsidEncoder As clsid)
   
   LSet m_clsIdEncoder = clsidEncoder
   
   Dim listSize As Long
   Dim status As GpStatus
   status = GdipGetEncoderParameterListSize( _
      cBitmap.nativeImage, clsidEncoder, listSize)
   If (status = NotImplemented) Then
      ' ok
   Else
      SetStatusHelper status
   End If
   
   m_lSize = listSize
   
   If (listSize > 0) Then
      ' allocate buffer:
      ReDim b(0 To listSize - 1) As Byte
      Dim lPtrBuff As Long
      lPtrBuff = VarPtr(b(0))
      SetStatusHelper LocalGdipGetEncoderParameterList( _
         cBitmap.nativeImage, _
         clsidEncoder, _
         listSize, _
         lPtrBuff)
      Dim p As EncoderParameters
      RtlMoveMemory p, b(0), Len(p)
      m_count = p.count
      If (m_count > 0) Then
         ReDim m_cParameters(1 To m_count) As GDIPEncoderParameter
         
         Dim i As Long
         Dim pP As EncoderParameter
         Dim lStart As Long
         lStart = 4
         
         For i = 1 To m_count
            RtlMoveMemory pP, b(lStart), Len(pP)
            Set m_cParameters(i) = New GDIPEncoderParameter
            m_cParameters(i).fInit pP.Guid, pP.NumberOfValues, pP.Type, pP.ValuePtr
            lStart = lStart + Len(pP)
         Next i
         
      End If
      
   End If

End Sub
