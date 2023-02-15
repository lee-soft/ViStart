VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_gfx As Long
Private m_lastResult As GpStatus

Private m_DrawStringRect As RECTF 'null layout seems to mean automatic
Private m_DrawStringFormat As Long

Friend Property Get nativeGraphics() As Long
   nativeGraphics = m_gfx
End Property

Public Function GetDpiX() As Single
    SetStatus GdipGetDpiX(m_gfx, GetDpiX)
End Function

Public Sub FromHDC(ByVal lhDC As Long)
   Dispose
   SetStatus GdipCreateFromHDC(lhDC, m_gfx), "FromHDC()"
End Sub

Public Sub FromHDCDevice(ByVal lhDC As Long, ByVal hDevice As Long)
   Dispose
   SetStatus GdipCreateFromHDC2(lhDC, hDevice, m_gfx), "FromHDCDevice()"
End Sub

Public Sub FromHWnd(ByVal lHWnd As Long)
   Dispose
   SetStatus GdipCreateFromHWND(lHWnd, m_gfx), "FromhWnd()"
End Sub

Public Sub FromImage(Image As GDIPImage)
   Dispose
   SetStatus GdipGetImageGraphicsContext(Image.nativeImage, m_gfx)
End Sub

Public Sub Flush(ByVal intention As FlushIntention)
   GdipFlush m_gfx, intention
End Sub

Public Function GetHDC() As Long
Dim lhDC As Long
   m_lastResult = GdipGetDC(m_gfx, lhDC)
   GetHDC = lhDC
End Function

Public Function ReleaseHDC(ByVal hdc As Long)
   m_lastResult = GdipReleaseDC(m_gfx, hdc)
End Function

Public Function SetRenderingOrigin(ByVal X As Long, ByVal Y As Long) As GpStatus
   SetRenderingOrigin = SetStatus( _
            GdipSetRenderingOrigin( _
                m_gfx, X, Y) _
            )
End Function
Public Function GetRenderingOrigin(ByRef X As Long, ByRef Y As Long) As GpStatus
   GetRenderingOrigin = SetStatus( _
            GdipGetRenderingOrigin( _
                m_gfx, X, Y) _
            )
End Function

Public Property Let CompositingMode(ByVal Value As CompositingMode)
   SetStatus GdipSetCompositingMode(m_gfx, Value)
End Property
Public Property Get CompositingMode() As CompositingMode
Dim mode As CompositingMode
   SetStatus GdipGetCompositingMode(m_gfx, mode)
   CompositingMode = mode
End Property

Public Property Let CompositingQuality(ByVal Value As CompositingQuality)
   SetStatus GdipSetCompositingQuality(m_gfx, Value)
End Property
Public Property Get CompositingQuality() As CompositingQuality
Dim Value As CompositingQuality
   SetStatus GdipGetCompositingQuality(m_gfx, Value)
   CompositingQuality = Value
End Property

Public Property Let TextRenderingHint(ByVal Value As TextRenderingHint)
   SetStatus GdipSetTextRenderingHint(m_gfx, Value)
End Property
Public Property Get TextRenderingHint() As TextRenderingHint
Dim Value As TextRenderingHint
   SetStatus GdipGetTextRenderingHint(m_gfx, Value)
   TextRenderingHint = Value
End Property

Public Property Let TextContrast(ByVal Value As Long)
   SetStatus GdipSetTextContrast(m_gfx, Value)
End Property
Public Property Get TextContrast() As Long
Dim Value As Long
   SetStatus GdipGetTextContrast(m_gfx, Value)
   TextContrast = Value
End Property

Public Property Let InterpolationMode(ByVal Value As InterpolationMode)
   SetStatus GdipSetInterpolationMode(m_gfx, Value)
End Property
Public Property Get InterpolationMode() As InterpolationMode
Dim Value As InterpolationMode
   SetStatus GdipGetInterpolationMode(m_gfx, Value)
   InterpolationMode = Value
End Property

Public Property Let SmoothingMode(ByVal Value As SmoothingMode)
   SetStatus GdipSetSmoothingMode(m_gfx, Value)
End Property
Public Property Get SmoothingMode() As SmoothingMode
Dim Value As SmoothingMode
   SetStatus GdipGetSmoothingMode(m_gfx, Value)
   SmoothingMode = Value
End Property

Public Property Let PixelOffsetMode(ByVal Value As PixelOffsetMode)
   SetStatus GdipSetPixelOffsetMode(m_gfx, Value)
End Property
Public Property Get PixelOffsetMode() As PixelOffsetMode
Dim Value As PixelOffsetMode
   SetStatus GdipGetPixelOffsetMode(m_gfx, Value)
   PixelOffsetMode = Value
End Property

Public Sub DrawString( _
      ByRef s As String, _
      ByRef Font As GDIPFont, _
      ByRef brush As GDIPBrush, _
      ByRef point As POINTF)
    
    m_DrawStringRect.Left = point.X
    m_DrawStringRect.Top = point.Y

    GdipDrawString m_gfx, s, -1, Font.Value, m_DrawStringRect, m_DrawStringFormat, brush.Value
End Sub

Public Sub DrawImagePointF( _
      Image As GDIPImage, _
      point As POINTF _
   )
   DrawImagePointFv Image, point.X, point.Y
End Sub
Public Sub DrawImagePointFv( _
      Image As GDIPImage, _
      ByVal X As Single, _
      ByVal Y As Single _
   )
   SetStatus GdipDrawImage(m_gfx, Image.nativeImage, X, Y)
End Sub
Public Sub DrawImageRectF( _
      Image As GDIPImage, _
      RECT As RECTF _
   )
   DrawImageRectFv Image, RECT.Left, RECT.Top, RECT.Width, RECT.Height
End Sub
Public Sub DrawImageRectFv( _
      Image As GDIPImage, _
      ByVal X As Single, _
      ByVal Y As Single, _
      ByVal Width As Single, _
      ByVal Height As Single _
   )
   SetStatus GdipDrawImageRect(m_gfx, _
         Image.nativeImage, _
         X, _
         Y, _
         Width, _
         Height)
End Sub
Public Sub DrawImagePointL( _
      Image As GDIPImage, _
      point As gdiplus.POINTL _
   )
   DrawImagePointLv Image, point.X, point.Y
End Sub
Public Sub DrawImagePointLv( _
      Image As GDIPImage, _
      ByVal X As Long, _
      ByVal Y As Long _
   )
   SetStatus GdipDrawImageI(m_gfx, _
         Image.nativeImage, _
         X, _
         Y)
End Sub

Public Sub DrawImageRectL( _
      Image As GDIPImage, _
      RECT As gdiplus.RECTL)
   DrawImageRectLv Image, RECT.Left, RECT.Top, RECT.Width, RECT.Height
End Sub

Public Sub DrawImageRectLv( _
      Image As GDIPImage, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal Width As Long, _
      ByVal Height As Long _
   )
   SetStatus GdipDrawImageRectI(m_gfx, _
      Image.nativeImage, _
      X, _
      Y, _
      Width, _
      Height)

End Sub
'
'    // Affine Draw Image
'    // destPoints.length = 3: rect => parallelogram
'    //     destPoints[0] <=> top-left corner of the source rectangle
'    //     destPoints[1] <=> top-right corner
'    //     destPoints[2] <=> bottom-left corner
'    // destPoints.length = 4: rect => quad
'    //     destPoints[3] <=> bottom-right corner
'
Public Sub DrawImageAffineF( _
      Image As GDIPImage, _
      destPoints() As POINTF, _
      ByVal count As Long _
   )
   If (count <> 3 And count <> 4) Then
      SetStatus InvalidParameter
   Else
      SetStatus GdipDrawImagePoints(m_gfx, _
         Image.nativeImage, _
         destPoints(0), count)
   End If
End Sub

Public Sub DrawImageAffineL( _
      Image As GDIPImage, _
      destPoints() As gdiplus.POINTL, _
      ByVal count As Long _
   )
   If (count <> 3 And count <> 4) Then
      SetStatus InvalidParameter
   Else
      SetStatus GdipDrawImagePointsI(m_gfx, _
         Image.nativeImage, _
         destPoints(0), count)
   End If

End Sub

Public Sub DrawImageStretchF( _
      Image As GDIPImage, _
      ByVal X As Single, _
      ByVal Y As Single, _
      ByVal srcx As Single, _
      ByVal srcy As Single, _
      ByVal srcwidth As Single, _
      ByVal srcheight As Single, _
      ByVal srcUnit As GpUnit)

   SetStatus GdipDrawImagePointRect(m_gfx, _
         Image.nativeImage, _
         X, Y, _
         srcx, srcy, _
         srcwidth, srcheight, srcUnit)

End Sub
Public Sub DrawImageStretchAttrF( _
      Image As GDIPImage, _
      destRect As RECTF, _
      ByVal srcx As Single, _
      ByVal srcy As Single, _
      ByVal srcwidth As Single, _
      ByVal srcheight As Single, _
      ByVal srcUnit As GpUnit, _
      ByVal imgAttributes As Long, _
      ByVal callback As Long, _
      ByVal callbackData As Long _
   )
   SetStatus GdipDrawImageRectRect(m_gfx, _
         Image.nativeImage, _
         destRect.Left, _
         destRect.Top, _
         destRect.Width, _
         destRect.Height, _
         srcx, srcy, _
         srcwidth, srcheight, _
         srcUnit, _
         imgAttributes, _
         callback, _
         callbackData)
End Sub

'public sub DrawImage(IN Image* image,
'                     IN const PointF* destPoints,
'                     IN INT count,
'                     IN REAL srcx,
'                     IN REAL srcy,
'                     IN REAL srcwidth,
'                     IN REAL srcheight,
'                     IN Unit srcUnit,
'                     IN const ImageAttributes* imageAttributes = NULL,
'                     IN DrawImageAbort callback = NULL,
'                     IN VOID* callbackData = NULL)
'    {
'        return SetStatus(DllExports::GdipDrawImagePointsRect(nativeGraphics,
'                                                             image ? image->nativeImage
'                                                                   : NULL,
'                                                             destPoints, count,
'                                                             srcx, srcy,
'                                                             srcwidth,
'                                                             srcheight,
'                                                             srcUnit,
'                                                             imageAttributes
'                                                              ? imageAttributes->nativeImageAttr
'                                                              : NULL,
'                                                             callback,
'                                                             callbackData));
'    }
'
Public Sub DrawImageStretchL( _
      Image As GDIPImage, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal srcx As Long, _
      ByVal srcy As Long, _
      ByVal srcwidth As Long, _
      ByVal srcheight As Long, _
      ByVal srcUnit As GpUnit _
   )
   SetStatus GdipDrawImagePointRectI(m_gfx, _
         Image.nativeImage, _
         X, Y, _
         srcx, srcy, _
         srcwidth, srcheight, _
         srcUnit)
End Sub

Public Sub DrawImageStretchAttrL( _
      Image As GDIPImage, _
      destRect As gdiplus.RECTL, _
      ByVal srcx As Long, _
      ByVal srcy As Long, _
      ByVal srcwidth As Long, _
      ByVal srcheight As Long, _
      ByVal srcUnit As GpUnit, _
      ByVal imgAttributes As Long, _
      ByVal callback As Long, _
      ByVal callbackData As Long _
   )
   SetStatus GdipDrawImageRectRectI(m_gfx, _
      Image.nativeImage, _
      destRect.Left, destRect.Top, _
      destRect.Width, destRect.Height, _
      srcx, srcy, srcwidth, srcheight, _
      srcUnit, _
      imgAttributes, _
      callback, _
      callbackData)

End Sub

Public Sub DrawImage(Image As GDIPImage, _
                        X As Single, _
                        Y As Single, _
                        Width As Single, _
                        Height As Single, _
                        Optional ImageAttributes As GDIPImageAttributes)
    
    If ImageAttributes Is Nothing Then
        SetStatus GdipDrawImageRect(m_gfx, Image.nativeImage, X, Y, Width, Height), "DrawImage"
    Else
        SetStatus GdipDrawImageRectRectI(m_gfx, Image.nativeImage, X, Y, Width, Height, 0, 0, Image.Width, Image.Height, UnitPixel, ImageAttributes.nativeAttributes, 0, 0), "DrawImage"
    End If
End Sub

Public Sub DrawImageRect(Image As GDIPImage, _
                          X As Long, _
                          Y As Long, _
                          Width As Long, _
                          Height As Long, _
                          srcx As Long, _
                          srcy As Long)

    
    SetStatus GdipDrawImageRectRect(m_gfx, Image.nativeImage, X, Y, Width, Height, srcx, srcy, Width, Height, UnitPixel)
End Sub

Public Sub DrawCachedBitmap(BITMAP As GDIPCachedBitmap, _
                        X As Long, _
                        Y As Long)
                        
    SetStatus GdipDrawCachedBitmap(m_gfx, BITMAP.nativeImage, X, Y)
End Sub

Public Sub DrawPath( _
    pen As GDIPPen, _
    GraphicsPath As GDIPGraphicPath _
)

    SetStatus GdipDrawPath(m_gfx, pen.Value, GraphicsPath.Value)
End Sub

Public Sub fillPath( _
    brush As GDIPBrush, _
    Path As GDIPGraphicPath _
)
    SetStatus GdipFillPath(m_gfx, brush.Value, Path.Value)
End Sub

Public Sub DrawArc( _
        pen As GDIPPen, _
        X As Single, _
        Y As Single, _
        Width As Single, _
        Height As Single, _
        startAngle As Single, _
        sweepAngle As Single)
    
    SetStatus GdipDrawArc(m_gfx, pen.Value, X, Y, Width, Height, startAngle, sweepAngle)
End Sub

Public Function MeasureString( _
        ByRef srcString As String, _
        ByRef srcFont As GDIPFont) As RECTF

Dim Layout As RECTF
Dim boundingBox As RECTF

Dim codeFitted As Long
Dim linesFilled As Long

    SetStatus GdipMeasureString(m_gfx, srcString, -1, srcFont.Value, _
                    Layout, m_DrawStringFormat, boundingBox, codeFitted, linesFilled)
                    
    MeasureString = boundingBox
End Function
        
   
'
'    Status DrawImage(IN Image* image,
'                     IN const Point* destPoints,
'                     IN INT count,
'                     IN INT srcx,
'                     IN INT srcy,
'                     IN INT srcwidth,
'                     IN INT srcheight,
'                     IN Unit srcUnit,
'                     IN const ImageAttributes* imageAttributes = NULL,
'                     IN DrawImageAbort callback = NULL,
'                     IN VOID* callbackData = NULL)
'    {
'        return SetStatus(DllExports::GdipDrawImagePointsRectI(nativeGraphics,
'                                                              image ? image->nativeImage
'                                                                    : NULL,
'                                                              destPoints,
'                                                              count,
'                                                              srcx,
'                                                              srcy,
'                                                              srcwidth,
'                                                              srcheight,
'                                                              srcUnit,
'                                                              imageAttributes
'                                                               ? imageAttributes->nativeImageAttr
'                                                               : NULL,
'                                                              callback,
'                                                              callbackData));
'    }


Public Sub Dispose()
    If Not (m_gfx = 0) Then
       GdipDeleteGraphics m_gfx
       GdipDeleteStringFormat m_DrawStringFormat
       
       m_gfx = 0
       m_DrawStringFormat = 0
    End If
End Sub

Public Function Clear(Optional ByRef r_Colour As Long = -1) As GpStatus
    If (r_Colour = -1) Then
        '0 is Black
        Clear = GdipGraphicsClear(m_gfx, 0)
    Else
        Clear = GdipGraphicsClear(m_gfx, GetRGB_VB2GDIP(r_Colour))
    End If
End Function

Private Function SetStatus(ByVal status As GpStatus, Optional ByRef szRoutine As String) As GpStatus
   m_lastResult = SetStatusHelper(status, "GDIPGraphics::" & szRoutine)
End Function

Private Sub Class_Initialize()
    GdipCreateStringFormat 0, 0, m_DrawStringFormat
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub
