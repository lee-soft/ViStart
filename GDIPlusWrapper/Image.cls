VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LocalGdipSaveImageToFile Lib "gdiplus.dll" Alias "GdipSaveImageToFile" _
   (ByVal Image As Long, _
   ByVal FileName As Long, _
   clsidEncoder As CLSID, _
   ByVal lPtrBuff As Long) As Long

Private Declare Function LocalGdipGetPropertyItem Lib "gdiplus.dll" Alias "GdipGetPropertyItem" _
   (ByVal img As Long, _
   ByVal lId As Long, _
   ByVal lSize As Long, _
   ByVal lPtrBuff As Long) As Long
   
Private Declare Function GdipAlloc Lib "gdiplus.dll" (ByVal Size As Long) As Long
Private Declare Function GdipFree Lib "gdiplus.dll" (ByVal ptr As Long) As Long
   
Private m_bUseEmbeddedColorManagement As Boolean
Private m_img As Long
Private m_lastResult As GpStatus

Private m_cEncoderParams() As GDIPEncoderParameterList
Private m_iEncoderParameterListCount As Long

Friend Sub fInit(ByVal nativeImage As Long, ByVal lastResult)
   Dispose
   m_img = nativeImage
   m_lastResult = lastResult
End Sub

Friend Property Get nativeImage() As Long
   nativeImage = m_img
End Property

Private Sub SetStatus(ByVal status As GpStatus)
   m_lastResult = SetStatusHelper(status)
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       FromBinary
' Description:       FromBinary method creates a new Image object based on
'                    binary data
' Parameters :       pngData() (Byte)
'--------------------------------------------------------------------------------
Public Sub FromBinary(ByRef pngData() As Byte)

    Dim oStream As IUnknown

    Dim nBytes  As Long

    Dim hMem    As Long

    Dim lpMem   As Long

    nBytes = UBound(pngData) - LBound(pngData) + 1
    ' Create and lock a memory buffer for image bytes.
    hMem = GlobalAlloc(GMEM_MOVEABLE, nBytes)

    If hMem Then
        lpMem = GlobalLock(hMem)

        If lpMem Then
            ' Copy image bytes to memory buffer, and unlock.
            Call mGDIPlus.CopyMemory(ByVal lpMem, pngData(LBound(pngData)), nBytes)
            Call GlobalUnlock(hMem)

            ' Create an IStream object in global memory buffer.
            If (CreateStreamOnHGlobal(hMem, False, oStream) = S_OK) Then
                
                SetStatus GdipLoadImageFromStream(ObjPtr(oStream), m_img)
            End If
        End If

        ' Release global memory object.
        Call GlobalFree(hMem)
    End If

End Sub


'--------------------------------------------------------------------------------
' Procedure  :       FromResource
' Description:       FromResource method creates a Image object from the
'                    final binary's resource
' Parameters :       resourceID (String)
'                    resourceType (String)
'--------------------------------------------------------------------------------
Public Sub FromResource(ByVal resourceID As String, ByVal resourceType As String)

    Dim pngData() As Byte

    pngData = LoadResData(resourceID, resourceType)
    FromBinary pngData

End Sub

Public Sub FromFile(ByVal sFile As String)
   Dispose
   If (m_bUseEmbeddedColorManagement) Then
      SetStatus GdipLoadImageFromFileICM( _
            sFile, _
            m_img _
        )
   Else
      SetStatus GdipLoadImageFromFile( _
            sFile, _
            m_img _
        )
   End If
End Sub

Public Function Clone() As GDIPImage
Dim cloneimage As Long

   SetStatus GdipCloneImage(m_img, cloneimage)
   If Not (cloneimage = 0) Then
      Dim cloneR As New GDIPImage
      cloneR.fInit cloneimage, m_lastResult
      Set Clone = cloneR
   End If

End Function

Public Property Get EncoderParameterList( _
      clsidEncoder As CLSID _
   )
Dim i As Long
Dim iFoundIndex As Long
   ' see if we need to add this item:
   For i = 1 To m_iEncoderParameterListCount
      If matchCLSID(m_cEncoderParams(i).clsidEncoder, clsidEncoder) Then
         iFoundIndex = i
         Exit For
      End If
   Next i
   If (iFoundIndex = 0) Then
      m_iEncoderParameterListCount = m_iEncoderParameterListCount + 1
      ReDim Preserve m_cEncoderParams(1 To m_iEncoderParameterListCount) As GDIPEncoderParameterList
      Set m_cEncoderParams(m_iEncoderParameterListCount) = New GDIPEncoderParameterList
      m_cEncoderParams(m_iEncoderParameterListCount).fInit Me, clsidEncoder
      iFoundIndex = m_iEncoderParameterListCount
   End If
   Set EncoderParameterList = m_cEncoderParams(iFoundIndex)
End Property

Private Function matchCLSID( _
      clsId1 As CLSID, _
      clsId2 As CLSID _
   ) As Boolean
Dim i As Long
   If (clsId1.Data1 = clsId2.Data1) Then
      If (clsId1.Data2 = clsId2.Data2) Then
         If (clsId1.Data3 = clsId2.Data3) Then
            For i = 0 To 7
               If Not (clsId1.Data4(i) = clsId2.Data4(i)) Then
                  Exit Function
               End If
            Next i
            matchCLSID = True
         End If
      End If
   End If

End Function
      
Public Function Save(ByVal FileName As String, clsidEncoder As CLSID)
    SetStatus LocalGdipSaveImageToFile(m_img, StrPtr(FileName), clsidEncoder, ByVal 0&)
      
End Function


Public Property Get ImgType() As ImageType
Dim itype  As ImageType
   itype = ImageTypeUnknown
   SetStatus GdipGetImageType(m_img, itype)
   ImgType = itype

End Property

Public Property Get PhysicalDimension() As SIZEF
Dim Size As SIZEF

   SetStatus GdipGetImageDimension(m_img, _
         Size.Width, Size.Height)
         
End Property

Public Property Get Width() As Long
Dim lWidth As Long
   
   SetStatus GdipGetImageWidth(m_img, lWidth)
   Width = lWidth
   
End Property

Public Property Get Height() As Long
Dim lHeight As Long
   
   SetStatus GdipGetImageHeight(m_img, lHeight)
   Height = lHeight

End Property

Public Property Get HorizontalResolution() As Single
Dim resolution As Single

   SetStatus GdipGetImageHorizontalResolution(m_img, resolution)
   HorizontalResolution = resolution

End Property

Public Property Get VerticalResolution() As Single
Dim resolution As Single

   SetStatus GdipGetImageVerticalResolution(m_img, resolution)
   VerticalResolution = resolution

End Property

Public Property Get Flags() As Long
    Dim lFlags As Long
    SetStatus GdipGetImageFlags(m_img, lFlags)
    Flags = lFlags
End Property

Public Property Get RawFormat() As CLSID
Dim format As CLSID
   SetStatus GdipGetImageRawFormat(m_img, format)
   RawFormat = format
End Property

Public Property Get ImgPixelFormat() As Long
Dim format As Long ' Should be PixelFormat

   SetStatus GdipGetImagePixelFormat(m_img, format)
   ImgPixelFormat = format

End Property

Public Property Get PaletteSize() As Long
Dim Size As Long
   SetStatus GdipGetImagePaletteSize(m_img, Size)
   PaletteSize = Size
End Property

Public Property Get Palette(ByVal Size As Long) As ColorPalette
Dim pal As ColorPalette
   SetStatus GdipGetImagePalette(m_img, pal, Size)
   Palette = pal
End Property

Public Sub SetPalette( _
      pal As ColorPalette _
   )
   SetStatus GdipSetImagePalette(m_img, pal)
End Sub

Public Function GetThumbnailImage( _
    ByVal thumbWidth As Long, _
    ByVal thumbHeight As Long, _
    ByVal callback As Long, _
    ByVal callbackData As Long _
   ) As GDIPImage
   Dim thumbimage As Long

   SetStatus GdipGetImageThumbnail(m_img, _
         thumbWidth, thumbHeight, _
         thumbimage, _
         callback, callbackData)

    Dim newImage As New GDIPImage
    newImage.fInit thumbimage, m_lastResult

    Set GetThumbnailImage = newImage

End Function


Public Property Get FrameDimensionsList() As GDIPFrameDimensionsList
   Dim cDimensionList As New GDIPFrameDimensionsList
   cDimensionList.fInit m_img
   Set FrameDimensionsList = cDimensionList
End Property


Public Sub SelectActiveFrame( _
    dimensionId As CLSID, _
    ByVal frameIndex As Long _
    )
   SetStatus GdipImageSelectActiveFrame(m_img, _
      dimensionId, _
      frameIndex)

End Sub
Public Sub RotateFlip( _
    ByVal rotFlipType As RotateFlipType _
    )
   SetStatus GdipImageRotateFlip(m_img, rotFlipType)

End Sub

Public Property Get PropertyCount() As Long
Dim numProperty As Long
   SetStatus GdipGetPropertyCount(m_img, _
      numProperty)
   PropertyCount = numProperty
End Property

Public Property Get PropertyItemForID(ByVal lId As Long) As GDIPPropertyItem
Dim lSize As Long
   SetStatus GdipGetPropertyItemSize(m_img, lId, lSize)
   If (lSize > 0) Then
      ReDim b(0 To lSize - 1) As Byte
      Dim lPtrBuff As Long
      lPtrBuff = VarPtr(b(0))
      SetStatus LocalGdipGetPropertyItem(m_img, lId, lSize, lPtrBuff)
      Dim p As gdiplus.PropertyItem
      Dim cItem As New GDIPPropertyItem
      Dim lDataSize As Long
      If Not (lPtrBuff = 0) And (lSize > 0) Then
         RtlMoveMemory p, ByVal lPtrBuff, Len(p)
         cItem.fInit p.Id, p.Length, p.Type, p.ValuePtr, lSize
      End If
      Set PropertyItemForID = cItem
   End If
End Property

Public Property Get PropertyItem(ByVal index As Long) As GDIPPropertyItem
Dim lCount As Long
   lCount = PropertyCount
   If (index > 0) And (index <= lCount) Then
      ' Get all property items :
      ReDim lPropId(0 To lCount - 1) As Long
      
      Dim lPtrList As Long
      lPtrList = VarPtr(lPropId(0))
      SetStatus GdipGetPropertyIdList(m_img, lCount, lPtrList)
      
      Set PropertyItem = PropertyItemForID(lPropId(index - 1))
   Else
      SetStatus InvalidParameter
   End If
End Property

Public Sub RemovePropertyItem(ByVal propId As Long)
   SetStatus GdipRemovePropertyItem(m_img, propId)
End Sub

Public Sub SetPropertyItem(Item As GDIPPropertyItem)
   Dim p As gdiplus.PropertyItem
   p.Id = Item.Id
   p.Length = Item.Length
   p.Type = Item.ItemType
   ReDim b(0 To Item.DataBufferSize - 1) As Byte
   Item.GetData b()
   p.ValuePtr = VarPtr(b(0))
   SetStatus GdipSetPropertyItem(m_img, p)
End Sub

Public Sub Dispose()
   If Not (m_img = 0) Then
      GdipDisposeImage m_img
      m_img = 0
   End If
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub
