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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_gfx As Long
Private m_lastResult As GpStatus

Private m_DrawStringRect   As RECTF 'null layout seems to mean automatic

Private m_DrawStringFormat As Long

Friend Property Get nativeGraphics() As Long
   nativeGraphics = m_gfx
End Property

Public Sub FromHDC(ByVal lhDC As Long)
   Dispose
   SetStatus GdipCreateFromHDC(lhDC, m_gfx)
End Sub

Public Sub FromHDCDevice(ByVal lhDC As Long, ByVal hDevice As Long)
   Dispose
   SetStatus GdipCreateFromHDC2(lhDC, hDevice, m_gfx)
End Sub

Public Sub FromHWnd(ByVal lHWnd As Long)
   Dispose
   SetStatus GdipCreateFromHWND(lHWnd, m_gfx)
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


'--------------------------------------------------------------------------------
' Procedure  :       Exclude
' Description:       The Exclude method updates this region to the portion of
'                    itself that does not intersect another region.
' Parameters :       theRect (gdiplus.RECTL)
'--------------------------------------------------------------------------------
Public Function Exclude(ByRef theRect As RECTF)
    SetStatus GdipSetClipRect(m_gfx, theRect.Left, theRect.Top, theRect.Width, theRect.Height, CombineModeExclude)
End Function


'--------------------------------------------------------------------------------
' Procedure  :       ResetExclusions
' Description:       The ResetExclusions method sets the clipping region of
'                    this Graphics object to an infinite region.
' Parameters :
'--------------------------------------------------------------------------------
Public Function ResetExclusions()
    SetStatus GdipResetClip(m_gfx)
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

'--------------------------------------------------------------------------------
' Procedure  :       DrawString
' Description:       The DrawString method draws a string based on a font and
'                    an origin for the string.
' Parameters :       s (String)
'                    Font (GDIPFont)
'                    Brush (GDIPBrush)
'                    point (POINTF)
'--------------------------------------------------------------------------------
Public Sub DrawString(ByRef s As String, _
                      ByRef Font As GDIPFont, _
                      ByRef Brush As GDIPBrush, _
                      ByRef point As POINTF)
    
    m_DrawStringRect.Left = point.X
    m_DrawStringRect.Top = point.Y

    SetStatus GdipDrawString(m_gfx, s, -1, Font.Value, m_DrawStringRect, _
        m_DrawStringFormat, Brush.Value)
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
      point As POINTL _
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
      ByVal srcX As Single, _
      ByVal srcY As Single, _
      ByVal srcwidth As Single, _
      ByVal srcheight As Single, _
      ByVal srcUnit As GpUnit)

   SetStatus GdipDrawImagePointRect(m_gfx, _
         Image.nativeImage, _
         X, Y, _
         srcX, srcY, _
         srcwidth, srcheight, srcUnit)

End Sub
Public Sub DrawImageStretchAttrF( _
      Image As GDIPImage, _
      destRect As RECTF, _
      ByVal srcX As Single, _
      ByVal srcY As Single, _
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
         srcX, srcY, _
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
      ByVal srcX As Long, _
      ByVal srcY As Long, _
      ByVal srcwidth As Long, _
      ByVal srcheight As Long, _
      ByVal srcUnit As GpUnit _
   )
   SetStatus GdipDrawImagePointRectI(m_gfx, _
         Image.nativeImage, _
         X, Y, _
         srcX, srcY, _
         srcwidth, srcheight, _
         srcUnit)
End Sub

Public Sub DrawImageStretchAttrL( _
      Image As GDIPImage, _
      destRect As gdiplus.RECTL, _
      ByVal srcX As Long, _
      ByVal srcY As Long, _
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
      srcX, srcY, srcwidth, srcheight, _
      srcUnit, _
      imgAttributes, _
      callback, _
      callbackData)

End Sub
   

'--------------------------------------------------------------------------------
' Procedure  :       DrawImage
' Description:       The DrawImage method draws an image.
' Parameters :       Image (GDIPImage)
'                    X (Single)
'                    y (Single)
'                    Width (Single)
'                    Height (Single)
'--------------------------------------------------------------------------------
Public Sub DrawImage(Image As GDIPImage, _
                      X As Single, _
                      Y As Single, _
                      Width As Single, _
                      Height As Single)
                        
    Debug.Print X, Y, Width, Height
    SetStatus GdipDrawImageRect(m_gfx, Image.nativeImage, X, Y, (Width), (Height))
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       DrawImageRect
' Description:       The DrawImageRect method draws an image.
' Parameters :       Image (GDIPImage)
'                    X (Long)
'                    y (Long)
'                    Width (Long)
'                    Height (Long)
'                    srcX (Long)
'                    srcY (Long)
'--------------------------------------------------------------------------------
Public Sub DrawImageRect(Image As GDIPImage, _
                         X As Long, _
                         Y As Long, _
                         Width As Long, _
                         Height As Long, _
                         srcX As Long, _
                         srcY As Long)
    
    SetStatus GdipDrawImageRectRect(m_gfx, Image.nativeImage, X, Y, Width, _
                Height, srcX, srcY, Width, Height, UnitPixel)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       DrawCachedBitmap
' Description:       The DrawCachedBitmap method draws the image stored in a
'                    CachedBitmap object.
' Parameters :       Bitmap (GDIPCachedBitmap)
'                    X (Long)
'                    y (Long)
'--------------------------------------------------------------------------------
Public Sub DrawCachedBitmap(BITMAP As GDIPCachedBitmap, X As Long, Y As Long)
                        
    SetStatus GdipDrawCachedBitmap(m_gfx, BITMAP.nativeImage, X, Y)
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       DrawPath
' Description:       The DrawPath method draws a sequence of lines and curves
'                    defined by a GraphicsPath object.
' Parameters :       Pen (GDIPPen)
'                    GraphicsPath (GDIPGraphicPath)
'--------------------------------------------------------------------------------
Public Sub DrawPath(Pen As GDIPPen, GraphicsPath As GDIPGraphicPath)

    SetStatus GdipDrawPath(m_gfx, Pen.Value, GraphicsPath.Value)
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       FillPath
' Description:       The FillPath method uses a brush to fill the interior of
'                    a path. If a figure in the path is not closed, this method
'                    treats the nonclosed figure as if it were closed by a
'                    straight line that connects the figure's starting and
'                    ending points.
' Parameters :       Brush (GDIPBrush)
'                    Path (GDIPGraphicPath)
'--------------------------------------------------------------------------------
Public Sub FillPath(Brush As GDIPBrush, Path As GDIPGraphicPath)
    SetStatus GdipFillPath(m_gfx, Brush.Value, Path.Value)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       FillRectangle
' Description:       The Graphics::FillRectangle method uses a brush to fill
'                    the interior of a rectangle.
' Parameters :       Brush (GDIPBrush)
'                    R (RECTF)
'--------------------------------------------------------------------------------
Public Sub FillRectangle(Brush As GDIPBrush, r As RECTF)
    
    SetStatus GdipFillRectangle(m_gfx, Brush.Value, r.Left, r.Top, r.Width, r.Height)
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       FillEllipse
' Description:       The FillEllipse method uses a brush to fill the interior
'                    of an ellipse that is specified by a rectangle.
' Parameters :       Brush (GDIPPathGradientBrush)
'                    R (RECTF)
'--------------------------------------------------------------------------------
Public Sub FillEllipse(Brush As GDIPPathGradientBrush, r As RECTF)

    SetStatus GdipFillEllipse(m_gfx, Brush.Value, r.Left, r.Top, r.Width, r.Height)
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       DrawRectangle
' Description:       The DrawRectangle method draws a rectangle.
' Parameters :       Pen (GDIPPen)
'                    X (Single)
'                    y (Single)
'                    Width (Single)
'                    Height (Single)
'--------------------------------------------------------------------------------
Public Sub DrawRectangle(Pen As GDIPPen, _
                         X As Single, _
                         Y As Single, _
                         Width As Single, _
                         Height As Single)
        
    SetStatus GdipDrawRectangle(m_gfx, Pen.Value, X, Y, Width, Height)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       DrawArc
' Description:       DrawArc method draws an arc. The arc is part of an ellipse.
' Parameters :       Pen (GDIPPen)
'                    X (Single)
'                    y (Single)
'                    Width (Single)
'                    Height (Single)
'                    startAngle (Single)
'                    sweepAngle (Single)
'--------------------------------------------------------------------------------
Public Sub DrawArc(Pen As GDIPPen, _
                   X As Single, _
                   Y As Single, _
                   Width As Single, _
                   Height As Single, _
                   startAngle As Single, _
                   sweepAngle As Single)
    
    SetStatus GdipDrawArc(m_gfx, Pen.Value, X, Y, Width, Height, startAngle, sweepAngle)
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       Transform (Untested?, TODO: Rename to SetWorldTransform)
' Description:       The Transform function sets a two-dimensional linear
'                    transformation between world space and page space for the
'                    specified device context. This transformation can be used
'                    to scale, rotate, shear, or translate graphics output.
' Parameters :       theMatrix (GDIPMatrix)
'--------------------------------------------------------------------------------
Public Function Transform(ByRef theMatrix As GDIPMatrix)
    GdipSetWorldTransform m_gfx, theMatrix.Value
End Function

'--------------------------------------------------------------------------------
' Procedure  :       MeasureStringEx
' Description:       Another variation on MeasureString, does the same with
'                    different parameters
' Parameters :       srcString (String)
'                    srcFont (GDIPFont)
'                    linesFilled (Long)
'                    layout (RECTF)
'--------------------------------------------------------------------------------
Public Function MeasureStringEx(ByRef srcString As String, _
                                ByRef srcFont As GDIPFont, _
                                ByRef linesFilled As Long, _
                                ByRef Layout As RECTF) As RECTF
        
    Dim BoundingBox As RECTF

    Dim codeFitted  As Long

    SetStatus GdipMeasureString(m_gfx, srcString, -1, srcFont.Value, Layout, _
           m_DrawStringFormat, BoundingBox, codeFitted, linesFilled)
                    
    MeasureStringEx = BoundingBox
        
End Function

'--------------------------------------------------------------------------------
' Procedure  :       DotNet_MeasureString
' Description:       Same as MeasureStringEx but with different parameters
' Parameters :       szText (String)
'                    theFont (GDIPFont)
'                    layoutArea (SIZEF)
'                    charactersFitted (Long)
'                    linedFilled (Long)
'--------------------------------------------------------------------------------
Public Function DotNet_MeasureString(ByRef szText As String, _
                                     ByRef theFont As GDIPFont, _
                                     ByRef layoutArea As SIZEF, _
                                     Optional ByRef charactersFitted As Long, _
                                     Optional ByRef linedFilled As Long) As SIZEF

    Dim Layout      As RECTF

    Dim BoundingBox As RECTF

    Layout.Width = layoutArea.Width
    Layout.Height = layoutArea.Height

    SetStatus GdipMeasureString(m_gfx, szText, -1, theFont.Value, Layout, m_DrawStringFormat, BoundingBox, charactersFitted, linedFilled)
                  
    DotNet_MeasureString.Width = BoundingBox.Width
    DotNet_MeasureString.Height = BoundingBox.Height
End Function

'--------------------------------------------------------------------------------
' Procedure  :       MeasureString
' Description:       The MeasureString method measures the extent of
'                    the string in the specified font and layout rectangle.
' Parameters :       srcString (String)
'                    srcFont (GDIPFont)
'--------------------------------------------------------------------------------
Public Function MeasureString(ByRef srcString As String, _
                              ByRef srcFont As GDIPFont) As RECTF

    Dim Layout      As RECTF

    Dim BoundingBox As RECTF

    Dim codeFitted  As Long

    Dim linesFilled As Long

    SetStatus GdipMeasureString(m_gfx, srcString, -1, srcFont.Value, Layout, m_DrawStringFormat, BoundingBox, codeFitted, linesFilled)

    MeasureString = BoundingBox

End Function


'--------------------------------------------------------------------------------
' Procedure  :       MeasureStringWidth (TODO: Rename to GetMeasCharRangeCount)
' Source:            https://msdn.microsoft.com/en-us/library/windows/desktop/ms534720(v=vs.85).aspx
' Parameters :       szText (String)
'                    theFont (GDIPFont)
'--------------------------------------------------------------------------------
Public Function MeasureStringWidth(ByVal szText As String, _
                                   ByRef theFont As GDIPFont) As Single


'The following example defines three ranges of character positions within a
'string and sets those ranges in a StringFormat object. Next,
'the StringFormat::GetMeasurableCharacterRangeCount method is used to get the
'number of character ranges that are currently set in the StringFormat object.
'This number is then used to allocate a buffer large enough to store the
'regions that correspond with the ranges. Then, the MeasureCharacterRanges
'method is used to get the three regions of the display that are occupied by
'the characters that are specified by the ranges.

    Dim layoutRect    As RECTF

    Dim charRanges(0) As CharacterRange

    ' // Layout rectangle used for drawing string
    layoutRect.Left = 0: layoutRect.Top = 0: layoutRect.Width = 0: layoutRect.Height = 0

    charRanges(0).First = 0: charRanges(0).length = Len(szText)
   
    ' // Other variables
    Dim rgCharRangeRegions() As Long

    Dim i                    As Long                        ' // loop counter

    Dim count                As Long                    ' // number of character ranges set

    Dim theBounds            As RECTF
   
    ' // Set three ranges of character positions.
    SetStatus GdipSetStringFormatMeasurableCharacterRanges(m_DrawStringFormat, 1, charRanges(0))
    
    ' // Get the number of ranges that have been set, and allocate memory to
    ' // store the regions that correspond to the ranges.
    SetStatus GdipGetStringFormatMeasurableCharacterRangeCount(m_DrawStringFormat, count)

    If count = 0 Then
        'LogError 0, "MeasureStringWidth", "GDIPGraphics", "no character" & " " & szText & " " & theFont.Value & " " & Me.nativeGraphics

        Exit Function

    End If

    ReDim rgCharRangeRegions(count - 1)

    For i = 0 To count - 1
        SetStatus GdipCreateRegion(rgCharRangeRegions(i))
    Next
   
    ' // Get the regions that correspond to the ranges within the string.
    ' // Then draw the string and show the regions.
    SetStatus GdipMeasureCharacterRanges(m_gfx, szText, -1, theFont.Value, layoutRect, m_DrawStringFormat, count, rgCharRangeRegions(0))
       
    SetStatus GdipGetRegionBounds(rgCharRangeRegions(0), m_gfx, theBounds)
    
    For i = 0 To count - 1
        If rgCharRangeRegions(i) Then GdipDeleteRegion rgCharRangeRegions(i)
    Next
    
    MeasureStringWidth = theBounds.Width

End Function
'-
   
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

'--------------------------------------------------------------------------------
' Procedure  :       Clear
' Description:       Clear method clears a Graphics object to a specified color.
' Parameters :       r_Colour (Long = -1)
'--------------------------------------------------------------------------------
Public Function Clear(Optional ByRef r_Colour As Long = -1) As GpStatus

    If (r_Colour = -1) Then
        '0 is Black
        Clear = GdipGraphicsClear(m_gfx, 0)
    Else
        Clear = GdipGraphicsClear(m_gfx, GetRGB_VB2GDIP(r_Colour))
    End If

End Function

Private Function SetStatus(ByVal status As GpStatus) As GpStatus
   m_lastResult = SetStatusHelper(status)
End Function

Private Sub Class_Initialize()
    GdipStringFormatGetGenericTypographic m_DrawStringFormat
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub
