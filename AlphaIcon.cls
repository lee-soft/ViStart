VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlphaIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Icon As GDIPBitmap
Private m_IconXP As GDIPBitmap

Private m_IsAlphaBitmap As Boolean

Public Property Get Image() As GDIPImage
    Set Image = m_IconXP.Image
End Property

Public Function CreateFromHICON(ByVal icoHandle As Long)

Dim ii As ICONINFO

Dim bmData As BitmapData
Dim bmBounds As gdiplus.RECTL

Dim X As Long
Dim Y As Long

    Set m_Icon = New GDIPBitmap
    Set m_IconXP = New GDIPBitmap

    If GetIconInfo(icoHandle, ii) = 0 Then
        Debug.Print "Error retrieving icon info!"
    End If
    
    m_Icon.CreateFromHBITMAP ii.hbmColor, 0
    
    DeleteObject ii.hbmColor
    DeleteObject ii.hbmMask
        
    If m_Icon.Image.ImgPixelFormat = PixelFormat.Format32bppArgb Then
    End If
        

        'If BITMAP.GetPixelFormatSize(m_Icon.PixelFormat) < 32 Then
            'Return ico.ToBitmap
        'End If
        
    bmBounds.Width = m_Icon.Image.Width
    bmBounds.Height = m_Icon.Image.Height

    bmData = m_Icon.LockBits(bmBounds, _
                ImageLockModeRead, _
                m_Icon.Image.ImgPixelFormat)

    m_IconXP.CreateFromSizeFormatData bmData.Height, _
                bmData.Width, _
                bmData.stride, _
                Format32bppArgb, _
                bmData.Scan0Ptr


        
    m_IsAlphaBitmap = False

    For Y = 0 To bmData.Height - 1
        For X = 0 To bmData.Width - 1
            
            Dim PixelColor As ARGB
            Long2ARGB m_IconXP.GetPixel(X, Y), PixelColor
                'PixelColor = Color.FromArgb(Marshal.ReadInt32(bmData.Scan0, (bmData.Stride * y) + (4 * x)))
            If PixelColor.A > 0 And PixelColor.A < 255 Then
                m_IsAlphaBitmap = True
                Exit For
            End If
        Next
        If m_IsAlphaBitmap Then Exit For
    Next

    m_Icon.UnlockBits bmData
        
    If Not m_IsAlphaBitmap Then
        m_Icon.Dispose
        m_IconXP.Dispose
        
        m_IconXP.CreateFromHICON icoHandle
    End If
    
End Function

