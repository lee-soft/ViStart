Attribute VB_Name = "mGDIPlus"
Option Explicit

Private m_token As Long
Private gfx As Long
Private m_LastError As Long
Public g_IgnoreGDIErrors As Boolean

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Public Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Boolean, ppstm As Any) As Long

Public Function GetErrorStatus() As GpStatus
    GetErrorStatus = m_LastError
End Function

Public Function TestFunction()
    TestFunction = 2
End Function

Public Function GDIPlusCreate() As Boolean
Dim gpInput As GdiplusStartupInput
Dim token As Long
   gpInput.GdiplusVersion = 1
   If GdiplusStartup(token, gpInput) = Ok Then
      m_token = token
      GDIPlusCreate = True

   End If
End Function

Public Sub GDIPlusDispose()
   If Not (m_token = 0) Then
      GdiplusShutdown m_token
      m_token = 0
   End If
End Sub

Public Function PtrToString(ByVal lPtr As Long) As String
Dim lSize As Long
Dim b() As Byte
Dim s As String
   If Not (lPtr = 0) Then
      lSize = lstrlenW(lPtr)
      If ((lSize > 0) And (lSize < &H10000)) Then
         ReDim b(0 To (lSize * 2) - 1) As Byte
         RtlMoveMemory b(0), ByVal lPtr, lSize * 2
         s = b
      End If
   End If
   PtrToString = b
End Function

Public Function SetStatusHelper(ByVal status As GpStatus, Optional szRoutineName As String) As GpStatus
    If (status = Ok) Then
       'ok
    Else
       'Err.Raise 1048 + status, App.EXEName & ".GDIP", "GDI+ Error " & status
       Logger.Error 1048 + status & " " & App.EXEName & ".GDIP", szRoutineName
    End If
    
    m_LastError = status
    SetStatusHelper = status
End Function

Public Function GetGuidString(Guid As CLSID) As String
Dim i As Long
Dim sGuid As String

   sGuid = "{" & hexPad(Guid.Data1, 8) & "-" & hexPad(Guid.Data2, 4) & "-" & hexPad(Guid.Data3, 4) & "-"
   sGuid = sGuid & hexPad(Guid.Data4(0), 2) & hexPad(Guid.Data4(1), 2) & "-"
   For i = 2 To 7
      sGuid = sGuid & hexPad(Guid.Data4(i), 2)
   Next i
   sGuid = sGuid & "}"
   GetGuidString = sGuid

End Function

Private Function hexPad(ByVal Value As Long, ByVal padSize As Long) As String
Dim sRet As String
Dim lMissing As Long
   sRet = Hex$(Value)
   lMissing = padSize - Len(sRet)
   If (lMissing > 0) Then
      sRet = String$(lMissing, "0") & sRet
   ElseIf (lMissing < 0) Then
      sRet = Mid$(sRet, -lMissing + 1)
   End If
   hexPad = sRet
End Function

Public Function UnsignedAdd(Start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
   
End Function

' Pass a VB/standard color to this function and get the GDI+ compatible color
Public Function GetRGB_VB2GDIP(ByVal lColor As Long, Optional ByVal Alpha As Byte = 255) As Long
   Dim rgbq As RGBQUAD
   CopyMemory rgbq, lColor, 4
   ' I must have done something wrong, but swapping Red and Blue works...
   GetRGB_VB2GDIP = ColorARGB(Alpha, rgbq.rgbBlue, rgbq.rgbGreen, rgbq.rgbRed)
End Function

' Use this in lieu of the Color class constructor
' Thanks to Richard Mason for help with this
Public Function ColorARGB(ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
   Dim bytestruct As COLORBYTES
   Dim result As COLORLONG
   
   With bytestruct
      .AlphaByte = Alpha
      .RedByte = Red
      .GreenByte = Green
      .BlueByte = Blue
   End With
   
   LSet result = bytestruct
   ColorARGB = result.longval
End Function


' Resize the picture using GDI plus
Sub gdipResize(graphics As Long, img As Long, hdc As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)

    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    'GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth img, OrWidth
        GdipGetImageHeight img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI graphics, img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI graphics, img, 0, 0, Width, Height
    End If
    'GdipDeleteGraphics Graphics
End Sub

Public Function CopyRectL(ByRef srcRect As gdiplus.RECTL, ByRef dstRect As gdiplus.RECTL)

    dstRect.Left = srcRect.Left
    dstRect.Top = srcRect.Top
    dstRect.Height = srcRect.Height
    dstRect.Width = srcRect.Width

End Function

Public Function CreateRectF(Left As Single, Top As Single, Height As Single, Width As Single) As gdiplus.RECTF

Dim newRectF As gdiplus.RECTF

    With newRectF
        .Left = Left
        .Top = Top
        .Height = Height
        .Width = Width
    End With
    
    CreateRectF = newRectF
End Function

Public Function CreateMatrix(V1 As Single, V2 As Single, V3 As Single, V4 As Single, V5 As Single, _
                             W1 As Single, W2 As Single, W3 As Single, W4 As Single, W5 As Single, _
                             X1 As Single, X2 As Single, X3 As Single, X4 As Single, X5 As Single, _
                             Y1 As Single, Y2 As Single, Y3 As Single, Y4 As Single, Y5 As Single, _
                             Z1 As Single, Z2 As Single, Z3 As Single, Z4 As Single, Z5 As Single) As ColorMatrix
                             
Dim clrMatrix As ColorMatrix

    clrMatrix.m(0, 0) = V1: clrMatrix.m(1, 0) = V2: clrMatrix.m(2, 0) = V3: clrMatrix.m(3, 0) = V4: clrMatrix.m(4, 0) = V5
    clrMatrix.m(0, 1) = W1: clrMatrix.m(1, 1) = W2: clrMatrix.m(2, 1) = W3: clrMatrix.m(3, 1) = W4: clrMatrix.m(4, 1) = W5
    clrMatrix.m(0, 2) = X1: clrMatrix.m(1, 2) = X2: clrMatrix.m(2, 2) = X3: clrMatrix.m(3, 2) = X4: clrMatrix.m(4, 2) = X5
    clrMatrix.m(0, 3) = Y1: clrMatrix.m(1, 3) = Y2: clrMatrix.m(2, 3) = Y3: clrMatrix.m(3, 3) = Y4: clrMatrix.m(4, 2) = Y5
    clrMatrix.m(0, 4) = Z1: clrMatrix.m(1, 4) = Z2: clrMatrix.m(2, 4) = Z3: clrMatrix.m(3, 4) = Z4: clrMatrix.m(4, 4) = Z5

    CreateMatrix = clrMatrix
End Function



