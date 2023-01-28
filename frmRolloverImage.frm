VERSION 5.00
Begin VB.Form frmRolloverImage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "ViStart_Png"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmRolloverImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ULW_ALPHA = &H2
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE As Long = -20

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Dim mDC As Long  ' Memory hDC
Dim mainBitmap As Long ' Memory Bitmap
Dim blendFunc32bpp As BLENDFUNCTION
Dim oldBitmap As Long
Dim mIndex As Long
Dim strMyId As String
Dim winSize As Size
Dim srcPoint As POINTL

Private m_Path As String

Public Event onMouseUp()

' Force the callback class to implement the interface defining the events.
Private mCallback As IPngImageEvents

Property Get Path() As String
    Path = m_Path
End Property

Property Let Path(ByVal strPath As String)
    m_Path = strPath
End Property

' Allow the callback object to be set. Very important.
Property Set callback(ByRef newObj As IPngImageEvents)
    Set mCallback = newObj
End Property

Property Get callback() As IPngImageEvents
    Set callback = mCallback
End Property

Property Let Id(ByVal strNewId As String)
    strMyId = strNewId
End Property

Property Get Id() As String
    Id = strMyId
End Property

Property Get index() As Long
    index = mIndex
End Property

Property Let index(new_index As Long)
    mIndex = new_index
End Property

Private Sub Form_Click()
    ' Raise an event, passing a parameter
    If (Not mCallback Is Nothing) Then _
        mCallback.onClick Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Raise an event, passing a parameter
    If (Not mCallback Is Nothing) Then _
        mCallback.onMouseDown Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmVistaMenu.Die
    
    RaiseEvent onMouseUp
    ' Raise an event, passing a parameter
    If (Not mCallback Is Nothing) Then _
        mCallback.onMouseUp Me, Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SelectObject mDC, oldBitmap

    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC mDC
End Sub

Function MakeTransFromImage(ByRef thePng As GDIPImage)

   Dim tempBI As BITMAPINFO
   Dim tempBlend As BLENDFUNCTION      ' Used to specify what kind of blend we want to perform
   Dim lngHeight As Long, lngWidth As Long
   Dim curWinLong As Long
   Dim newGraphics As GDIPGraphics
   
   Dim hBitmap As Long
   
   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32    ' Each pixel is 32 bit's wide
      .biHeight = Me.ScaleHeight  ' Height of the form
      .biWidth = Me.ScaleWidth    ' Width of the form
      .biPlanes = 1   ' Always set to 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
   End With
   
    mDC = CreateCompatibleDC(Me.hdc)
   
    If mDC = 0 Then
        MsgBox "CreateCompatibleDC Failed", vbCritical
        
        MakeTransFromImage = False
        Exit Function
    End If
   
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    
    If mainBitmap = 0 Then
        MsgBox "CreateDIBSection Failed", vbCritical
        
        MakeTransFromImage = False
        Exit Function
    End If
   
    oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected
    
    If oldBitmap = 0 Then
        MsgBox "SelectObject Failed", vbCritical
        
        MakeTransFromImage = False
        Exit Function
    End If
    
   Set newGraphics = New GDIPGraphics
   Call newGraphics.FromHDC(mDC)
   
   lngHeight = thePng.Height
   lngWidth = thePng.Width
   newGraphics.DrawImage thePng, 0, 0, CSng(lngWidth), CSng(lngHeight)
   
   curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
   SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED Or WS_EX_TOOLWINDOW

   srcPoint.X = 0
   srcPoint.Y = 0
   winSize.cx = Me.ScaleWidth
   winSize.cy = Me.ScaleHeight
    
   With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA ' 32 bit
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = 255
   End With
    
   Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)

End Function

Function MakeTrans(ByVal pngPath As String) As Boolean

   Dim tempBI As BITMAPINFO
   Dim tempBlend As BLENDFUNCTION      ' Used to specify what kind of blend we want to perform
   Dim curWinLong As Long
   
   Dim hBitmap As Long
   
   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32    ' Each pixel is 32 bit's wide
      .biHeight = Me.ScaleHeight  ' Height of the form
      .biWidth = Me.ScaleWidth    ' Width of the form
      .biPlanes = 1   ' Always set to 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
   End With
   
    mDC = CreateCompatibleDC(Me.hdc)
   
    If mDC = 0 Then
        MsgBox "CreateCompatibleDC Failed", vbCritical
        
        MakeTrans = False
        Exit Function
    End If
   
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    
    If mainBitmap = 0 Then
        MsgBox "CreateDIBSection Failed", vbCritical
        
        MakeTrans = False
        Exit Function
    End If
   
    oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected
    
    If oldBitmap = 0 Then
        MsgBox "SelectObject Failed", vbCritical
        
        MakeTrans = False
        Exit Function
    End If
   
   Dim newGraphics As New GDIPGraphics
   Call newGraphics.FromHDC(mDC)
   
   Dim newImage As New GDIPImage
   newImage.FromFile pngPath

    newGraphics.DrawImage newImage, CSng(0), CSng(0), newImage.Width, newImage.Height

   curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
   SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED Or WS_EX_TOOLWINDOW

    ' Make the window a top-most window so we can always see the cool stuff
   'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
   'SetOwner Me.hWnd, frmStartMenuBase.hWnd
   
   ' Needed for updateLayeredWindow call
   srcPoint.X = 0
   srcPoint.Y = 0
   winSize.cx = Me.ScaleWidth
   winSize.cy = Me.ScaleHeight
    
   With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA ' 32 bit
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = 255
   End With
   
   Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)

End Function

Public Property Get Alpha() As Byte
    Alpha = blendFunc32bpp.SourceConstantAlpha
End Property

Public Property Let Alpha(bNewAlpha As Byte)
    blendFunc32bpp.SourceConstantAlpha = bNewAlpha
    UpdateLayeredWindow Me.hWnd, Me.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA
End Property

