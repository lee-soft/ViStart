VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISoftX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, ByVal hlpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetDCPenColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal colorref As Long) As Long
Private Declare Function SetDCBrushColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal colorref As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Long) As Long


Public Enum DrawConstants
    DT_BOTTOM = &H8
    DT_CENTER = &H1
    DT_LEFT = &H0
    DT_RIGHT = &H2
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
    DT_NOPREFIX = &H800
    DT_WORD_ELLIPSIS = &H40000
    DT_CALCRECT = &H400&
    DT_SINGLELINE = &H20
    DT_MODIFYSTRING = &H10000
End Enum

Private mBackBuffer As pcMemDC
Private mrBackBuffer As RECT

Private mHwnd As Long
Private mSrcDc As Long

Private mSrcWidth As Long
Private mSrcHeight As Long
Private mSrcRect As RECT

Private paNull As POINTL

Private m_TextColour As Long

'Convert stdFont into IFont (to grab the hFont)
Private mFontH As IFont

Public Property Let TextColour(newColour As Long)
    m_TextColour = newColour
    
    SetTextColor mBackBuffer.hdc, m_TextColour
End Property

Public Function DrawRectangle(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    DrawRectangle = Rectangle(mBackBuffer.hdc, X1, Y1, X2, Y2)
End Function

Public Function MoveTo(ByVal X As Long, ByVal Y As Long)
    MoveToEx mBackBuffer.hdc, X, Y, 0
End Function

Public Function DrawLineTo(ByVal X2 As Long, ByVal Y2 As Long)
    DrawLineTo = LineTo(mBackBuffer.hdc, X2, Y2)
End Function

Public Function SetBrushColor(lngColour As Long) As Long
    SetBrushColor = SetDCBrushColor(mBackBuffer.hdc, lngColour)
End Function

Public Function SetPenColor(lngColour As Long) As Long
    SetPenColor = SetDCPenColor(mBackBuffer.hdc, lngColour)
End Function

Public Property Get hdc() As Long
    hdc = mBackBuffer.hdc
End Property

Public Property Let Font(new_stdFont As StdFont)
    Set mFontH = new_stdFont

    'Update the HDC font
    SelectObject mBackBuffer.hdc, mFontH.hFont
End Property

Public Property Let hWnd(new_Hwnd As Long)
    
    mHwnd = new_Hwnd
    
    mSrcDc = GetWindowDC(mHwnd)
    GetWindowRect mHwnd, mSrcRect
    
    mSrcWidth = mSrcRect.Right - mSrcRect.Left
    mSrcHeight = mSrcRect.Bottom - mSrcRect.Top
    
    mBackBuffer.Height = mSrcHeight
    mBackBuffer.Width = mSrcWidth
    
    mrBackBuffer.Right = mSrcWidth
    mrBackBuffer.Bottom = mSrcHeight

End Property

Sub SetDimensionVars()
    If mBackBuffer Is Nothing Then
        Exit Sub
    End If

    GetWindowRect mHwnd, mSrcRect
    
    mSrcWidth = mSrcRect.Right - mSrcRect.Left
    mSrcHeight = mSrcRect.Bottom - mSrcRect.Top
    
    mBackBuffer.Height = mSrcHeight
    mBackBuffer.Width = mSrcWidth
    
    mrBackBuffer.Right = mSrcWidth
    mrBackBuffer.Bottom = mSrcHeight

    'Update the HDC font
    SelectObject mBackBuffer.hdc, mFontH.hFont
    SetTextColor mBackBuffer.hdc, m_TextColour
    
    'Reset the objects
    SelectObject mBackBuffer.hdc, GetStockObject(19)
    SelectObject mBackBuffer.hdc, GetStockObject(18)
End Sub

Public Sub OpenScene(Optional theBrush As GDIBrush)
    
    If theBrush Is Nothing Then
        FillRect mBackBuffer.hdc, mrBackBuffer, vbWhite '(0 seems to be white)
    Else
        FillRect mBackBuffer.hdc, mrBackBuffer, theBrush.Value
    End If
End Sub

Public Sub AddSpriteEX(ByRef hdcSprite As Long, dvDestPosition As POINTL, dvSrcPosition As POINTL, lngWidth As Long, lngHeight As Long)
                  'Where on Source
                  
    BitBlt mBackBuffer.hdc, dvDestPosition.X, dvDestPosition.Y, _
                    lngWidth, lngHeight, _
                    hdcSprite, dvSrcPosition.X, dvSrcPosition.Y, vbSrcCopy
End Sub

Public Sub AddSprite(ByRef dcSprite As pcMemDC, dvDestPosition As POINTL)
                  'Where on Source
                  
    BitBlt mBackBuffer.hdc, dvDestPosition.X, dvDestPosition.Y, _
                    dcSprite.Width, dcSprite.Height, _
                    dcSprite.hdc, 0, 0, vbSrcCopy
End Sub

Public Function GetTextRect(sText As String) As RECT
Dim recSize As RECT

    DrawTextW mBackBuffer.hdc, StrPtr(sText), -1, recSize, DT_CALCRECT Or DT_NOPREFIX
    GetTextRect = recSize
End Function

Public Function DrawText(sText As String, destRect As RECT, Optional wFormat As DrawConstants = DT_LEFT Or DT_MODIFYSTRING)
    DrawTextW mBackBuffer.hdc, StrPtr(sText), Len(sText), destRect, CLng(wFormat)
End Function

Public Sub PresentSceneEx(srcRect As RECT)
    BitBlt mSrcDc, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, mBackBuffer.hdc, srcRect.Left, srcRect.Top, vbSrcCopy
End Sub

Public Sub PresentScene()
    BitBlt mSrcDc, 0, 0, mSrcWidth, mSrcHeight, mBackBuffer.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub Class_Initialize()
    Set mBackBuffer = New pcMemDC
End Sub

Private Sub Class_Terminate()
    Set mBackBuffer = Nothing

End Sub
