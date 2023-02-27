VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_bUseEmbeddedColorManagement As Boolean
Private m_img As GDIPImage
Private m_lastResult As GpStatus

Friend Sub fInit(ByVal nativeImage As Long, ByVal lastResult)
   Dispose
   m_img.fInit nativeImage, lastResult
   m_lastResult = lastResult
End Sub

Friend Property Get nativeImage() As Long
   nativeImage = m_img.nativeImage
End Property

Private Sub SetStatus(ByVal status As GpStatus)
   m_lastResult = SetStatusHelper(status)
End Sub

Public Property Get Image() As GDIPImage
   Set Image = m_img
End Property

Public Sub FromBITMAPINFO( _
      gdiBitmapInfo As gdiplus.BITMAPINFO, _
      ByVal ptrBitmapData As Long _
   )
Dim bm As Long
   Dispose
   SetStatus GdipCreateBitmapFromGdiDib(gdiBitmapInfo, ptrBitmapData, bm)
   m_img.fInit bm, m_lastResult
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       UnlockBits
' Description:       Unlocks this Bitmap from system memory.
' Parameters :       LockedBitmapData (BitmapData)
'--------------------------------------------------------------------------------
Public Function UnlockBits(LockedBitmapData As BitmapData)
    SetStatus GdipBitmapUnlockBits(m_img.nativeImage, LockedBitmapData)
End Function


'--------------------------------------------------------------------------------
' Procedure  :       LockBits
' Description:       LockBits method locks a rectangular portion of this bitmap
'                    and provides a temporary buffer that you can use to read
'                    or write pixel data in a specified format. Any pixel data
'                    that you write to the buffer is copied to the Bitmap
'                    object when you call UnlockBits
' Parameters :       theRect (gdiplus.RECTL)
'                    theFlags (ImageLockMode)
'                    thePixelFormat (PixelFormat)
'--------------------------------------------------------------------------------
Public Function LockBits(theRect As gdiplus.RECTL, _
                         theFlags As ImageLockMode, _
                         thePixelFormat As PixelFormat) As BitmapData

    Dim bmpData As BitmapData

    SetStatus GdipBitmapLockBits(m_img.nativeImage, theRect, theFlags, thePixelFormat, bmpData)
    LockBits = bmpData
End Function

Public Sub CreateFromSizeFormatData( _
      ByVal Width As Long, _
      ByVal Height As Long, _
      ByVal stride As Long, _
      ByVal format As Long, _
      ByVal ptrBitmapData As Long _
   )
Dim bm As Long
   Dispose
   SetStatus GdipCreateBitmapFromScan0(Width, Height, stride, format, ByVal ptrBitmapData, bm)
   m_img.fInit bm, m_lastResult
End Sub

Public Sub CreateFromSize( _
      ByVal Width As Long, _
      ByVal Height As Long _
   )
   CreateFromSizeFormat Width, Height, 0 ' don't care
End Sub

Public Sub CreateFromSizeFormat( _
      ByVal Width As Long, _
      ByVal Height As Long, _
      ByVal format As Long _
   )
Dim bm As Long
   Dispose
   SetStatus GdipCreateBitmapFromScan0(Width, Height, 0, format, ByVal 0&, bm)
   m_img.fInit bm, m_lastResult
End Sub

Public Sub CreateFromGraphics( _
      ByVal Width As Long, _
      ByVal Height As Long, _
      gfx As GDIPGraphics _
   )
Dim bm As Long
   Dispose
   SetStatus GdipCreateBitmapFromGraphics(Width, Height, gfx.nativeGraphics, bm)
   m_img.fInit bm, m_lastResult
End Sub

Public Sub CreateFromHICON( _
      ByVal hIcon As Long _
   )
Dim bm As Long
   Dispose
   SetStatus GdipCreateBitmapFromHICON(hIcon, bm)
   m_img.fInit bm, m_lastResult
End Sub

Public Sub CreateFromHBITMAP( _
      ByVal hBmp As Long, _
      ByVal hPal As Long _
   )
Dim bm As Long
   Dispose
   SetStatus GdipCreateBitmapFromHBITMAP(hBmp, hPal, bm)
   m_img.fInit bm, m_lastResult
End Sub

Public Sub CreateFromResource( _
      ByVal hInstance As Long, _
      ByVal sBitmapName As String _
   )
Dim bm As Long
   Dispose
   SetStatus GdipCreateBitmapFromResource(hInstance, sBitmapName, bm)
   m_img.fInit bm, m_lastResult
End Sub

Public Sub CreateFromStdPicture( _
      sPic As IPicture _
   )
Dim hBmp As Long
Dim hPal As Long
   hBmp = sPic.Handle
   hPal = sPic.hPal
   Dispose
   CreateFromHBITMAP hBmp, hPal
End Sub

Public Property Get hBitmap( _
      ByVal colorBackground As Long _
    ) As Long
Dim hBmReturn As Long
   SetStatus GdipCreateHBITMAPFromBitmap(m_img.nativeImage, hBmReturn, colorBackground)
   hBitmap = hBmReturn
End Function

Public Property Get hIcon() As Long
Dim hIconReturn As Long
   SetStatus GdipCreateHICONFromBitmap( _
         m_img.nativeImage, _
         hIconReturn)
   hIcon = hIconReturn
End Property

Public Property Get Picture() As IPicture
   
End Property

Public Function CloneSizeL( _
      RECT As gdiplus.RECTL, _
      ByVal PixelFormat As Long _
   ) As GDIPBitmap
   Set CloneSizeL = CloneSizeLv(RECT.Left, RECT.Top, RECT.Width, RECT.Height, PixelFormat)
End Function

Public Function CloneSizeLv( _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal Width As Long, _
      ByVal Height As Long, _
      ByVal PixelFormat As Long _
   ) As GDIPBitmap
   Dim c As New GDIPBitmap
   Dim bmNew As Long
   SetStatus GdipCloneBitmapAreaI(X, Y, Width, Height, PixelFormat, m_img.nativeImage, bmNew)
   If (m_lastResult = Ok) Then
      c.fInit bmNew, m_lastResult
   End If
   Set CloneSizeLv = c
End Function

Public Function CloneSizeF( _
      RECT As RECTF, _
      ByVal PixelFormat As Long _
   ) As GDIPBitmap
   Set CloneSizeF = CloneSizeFv(RECT.Left, RECT.Top, RECT.Width, RECT.Height, PixelFormat)
End Function

Public Function CloneSizeFv( _
      ByVal X As Single, _
      ByVal Y As Single, _
      ByVal Width As Single, _
      ByVal Height As Single, _
      ByVal PixelFormat As Long _
   ) As GDIPBitmap
   Dim c As New GDIPBitmap
   Dim bmNew As Long
   SetStatus GdipCloneBitmapArea(X, Y, Width, Height, PixelFormat, m_img.nativeImage, bmNew)
   If (m_lastResult = Ok) Then
      c.fInit bmNew, m_lastResult
   End If
   Set CloneSizeFv = c
End Function

'inline status
'Bitmap::LockBits(
'    IN const Rect* rect,
'    IN UINT flags,
'    IN PixelFormat format,
'    OUT BitmapData * lockedBitmapData
')
'{
'    return SetStatus(DllExports::GdipBitmapLockBits(
'                                    static_cast<GpBitmap*>(nativeImage),
'                                    rect,
'                                    flags,
'                                    format,
'                                    lockedBitmapData));
'}
'
'inline status
'Bitmap::UnlockBits(
'    IN BitmapData* lockedBitmapData
'    )
'{
'    return SetStatus(DllExports::GdipBitmapUnlockBits(
'                                    static_cast<GpBitmap*>(nativeImage),
'                                    lockedBitmapData));
'}

Public Function GetPixel( _
      ByVal X As Long, _
      ByVal Y As Long _
   ) As Long
Dim ARGB As Long
   SetStatus GdipBitmapGetPixel(m_img.nativeImage, X, Y, ARGB)
   GetPixel = ARGB
End Function

Public Function SetPixel( _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal ARGB As Long _
   )
   SetStatus GdipBitmapSetPixel(m_img.nativeImage, X, Y, ARGB)
End Function

Public Sub SetResolution( _
      ByVal xdpi As Single, _
      ByVal ydpi As Single _
   )
   SetStatus GdipBitmapSetResolution(m_img.nativeImage, xdpi, ydpi)
End Sub

Public Sub Dispose()
   m_img.Dispose
End Sub

Private Sub Class_Initialize()
   Set m_img = New GDIPImage
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub


