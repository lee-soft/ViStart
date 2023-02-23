VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPCachedBitmap"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : GDIPCachedBitmap
'    Description: A CachedBitmap object stores a bitmap in a format that is
'    optimized for display on a particular device. To display a cached bitmap,
'    call the GDIPGraphics::DrawCachedBitmap method.
'
'    Based on Microsoft's CachedBitmap class   :
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534425(v=vs.85).aspx
'--------------------------------------------------------------------------------


Option Explicit

Private m_img        As GDIPImage

Private m_lastResult As GpStatus

Friend Property Get nativeImage() As Long
    nativeImage = m_img.nativeImage
End Property

Friend Sub fInit(ByVal nativeImage As Long, ByVal lastResult)
    'Dispose
    m_img.fInit nativeImage, lastResult
    m_lastResult = lastResult
End Sub

Private Sub SetStatus(ByVal status As GpStatus)
    m_lastResult = SetStatusHelper(status)
End Sub

Public Function Dispose()
    m_img.Dispose
End Function

Public Function Constructor(BITMAP As GDIPBitmap, graphics As GDIPGraphics)

    Dim hBitmap As Long

    SetStatus GdipCreateCachedBitmap(BITMAP.nativeImage, graphics.nativeGraphics, hBitmap)
    fInit hBitmap, m_lastResult
End Function

Private Sub Class_Initialize()
    Set m_img = New GDIPImage
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub
