Attribute VB_Name = "mGDIPlus"
Option Explicit

Private m_token As Long
Private gfx As Long

Private m_LastError               As Long

Private m_stringFormat_Generic    As Long

Private m_tokenGraphicsObject     As GDIPGraphics

Private m_defaultBrushYellow      As GDIPBrush

Private m_defaultBrushBlue        As GDIPBrush

Private m_defaultBrushBlack       As GDIPBrush

Private m_defaultBrushWhite       As GDIPBrush

Private m_defaultBrushTransparent As GDIPBrush

Private m_defaultSolidBlackPen    As GDIPPen

Public g_IgnoreGDIErrors          As Boolean

Public Declare Function GlobalAlloc _
                Lib "kernel32.dll" (ByVal wFlags As Long, _
                                    ByVal dwBytes As Long) As Long

Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Dest As Any, _
                                      Src As Any, _
                                      ByVal cb As Long) As Long

Public Declare Function CreateStreamOnHGlobal _
               Lib "ole32" (ByVal hGlobal As Long, _
                            ByVal fDeleteOnRelease As Boolean, _
                            ppstm As Any) As Long

Public Const LANG_NEUTRAL As Long = &H0

Public Const GMEM_MOVEABLE = 2

Public Const S_OK = 0

Public Brushes As Collection


'--------------------------------------------------------------------------------
' Procedure  :       CreateWebColour
' Description:       Takes a Web/Hex colour and returns a Colour object
' Parameters :       theWebColour (String)
'--------------------------------------------------------------------------------
Public Function CreateWebColour(ByVal theWebColour As String) As Colour
    
    Dim newColour As New Colour
    newColour.SetColourByHex theWebColour
    
    Set CreateWebColour = newColour
End Function

'--------------------------------------------------------------------------------
' Procedure  :       SolidBlackPen
' Description:       Promises to always return a solid back pen object
'                    TODO: Write a similar routine for all basic .NET english
'                    colours
'                    IE: Megenta DarkBrown etc
' Parameters :
'--------------------------------------------------------------------------------
Public Function SolidBlackPen() As GDIPPen

    If m_defaultSolidBlackPen Is Nothing Then

        Dim black As New Colour

        Set m_defaultSolidBlackPen = New GDIPPen
        black.SetColourByHex "000000"
        m_defaultSolidBlackPen.Constructor black, 1, 255
    End If

    Set SolidBlackPen = m_defaultSolidBlackPen
End Function

'--------------------------------------------------------------------------------
' Procedure  :       Custom_Brush
' Description:       Generates a GDIPBrush from a Colour object
' Parameters :       theColour (Colour)
'--------------------------------------------------------------------------------
Public Function Custom_Brush(ByVal theColour As Colour) As GDIPBrush

    'FFFCA1
    Dim m_defaultBrush As New GDIPBrush

    m_defaultBrush.Colour = theColour
    
    Set Custom_Brush = m_defaultBrush
End Function


'--------------------------------------------------------------------------------
' Procedure  :       Brushes_White
' Description:       Promises to return a White GDIPBrush
' Parameters :
'--------------------------------------------------------------------------------
Public Function Brushes_White() As GDIPBrush

    If m_defaultBrushWhite Is Nothing Then
        Set m_defaultBrushWhite = New GDIPBrush
        m_defaultBrushWhite.Colour = CreateColour(vbWhite)
    End If
    
    Set Brushes_White = m_defaultBrushWhite
End Function

Public Function Brushes_Transparent() As GDIPBrush

    If m_defaultBrushTransparent Is Nothing Then
        Set m_defaultBrushTransparent = New GDIPBrush
        m_defaultBrushTransparent.Colour.SetColourByHex "000000"
        m_defaultBrushTransparent.Colour.Alpha = 0
    End If
    
    Set Brushes_Transparent = m_defaultBrushTransparent
End Function

Public Function Brushes_Black() As GDIPBrush

    If m_defaultBrushBlack Is Nothing Then
        Set m_defaultBrushBlack = New GDIPBrush
        m_defaultBrushBlack.Colour.SetColourByHex "000000"
    End If
    
    Set Brushes_Black = m_defaultBrushBlack
End Function

Public Function Brushes_Yellow() As GDIPBrush

    If m_defaultBrushYellow Is Nothing Then
        Set m_defaultBrushYellow = New GDIPBrush
        m_defaultBrushYellow.Colour = CreateColour(vbYellow)
    End If
    
    Set Brushes_Yellow = m_defaultBrushYellow
End Function

Public Function Brushes_Blue() As GDIPBrush

    If m_defaultBrushBlue Is Nothing Then
        Set m_defaultBrushBlue = New GDIPBrush
        m_defaultBrushBlue.Colour = CreateColour(vbBlue)
    End If
    
    Set Brushes_Blue = m_defaultBrushBlue
End Function


'--------------------------------------------------------------------------------
' Procedure  :       CreateColour
' Description:       A conveniant way to generate basic GDI+ colours returning
'                    The colour class
' Parameters :       theColour (ColorConstants)
'--------------------------------------------------------------------------------
Public Function CreateColour(theColour As ColorConstants) As Colour

    Dim newColour As New Colour
    newColour.Value = theColour

    Set CreateColour = newColour
End Function

'--------------------------------------------------------------------------------
' Procedure  :       CreateFontFamily
' Description:       A conveniant way to quickly create font families
' Parameters :       szFontName (String)
'--------------------------------------------------------------------------------
Public Function CreateFontFamily(szFontName As String) As GDIPFontFamily

    Dim newFontFamily As New GDIPFontFamily

    newFontFamily.Constructor szFontName

    Set CreateFontFamily = newFontFamily
End Function

Public Function GetErrorStatus() As GpStatus
    GetErrorStatus = m_LastError
End Function


Public Function GDIPlusCreate(Optional suppressErrors As Boolean = False) As Boolean
    g_IgnoreGDIErrors = suppressErrors

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

Public Function SetStatusHelper(ByVal status As GpStatus) As GpStatus
   If (status = Ok) Then
      ' ok
   Else
      If Not g_IgnoreGDIErrors Then
      Err.Raise 1048 + status, App.EXEName & ".GDIP", "GDI+ Error " & status
   End If
       'If status <> 2 Then LogError status, "GDI+ Error", "GDI+Framework"
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

Private Function hexPad(ByVal value As Long, ByVal padSize As Long) As String
Dim sRet As String
Dim lMissing As Long
   sRet = Hex$(value)
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


'--------------------------------------------------------------------------------
' Procedure  :       GetRGB_VB2GDIP
' Description:       Pass a VB/standard color to this function and get the
'                    GDI+ compatible color
' Author     :       Richard Mason?
' Parameters :       lColor (Long)
'                    Alpha (Byte = 255)
'--------------------------------------------------------------------------------
Public Function GetRGB_VB2GDIP(ByVal lColor As Long, _
                               Optional ByVal Alpha As Byte = 255) As Long

    Dim rgbq As RGBQUAD

    CopyMemory rgbq, lColor, 4
    ' I must have done something wrong, but swapping Red and Blue works...
    GetRGB_VB2GDIP = ColorARGB(Alpha, rgbq.rgbBlue, rgbq.rgbGreen, rgbq.rgbRed)
End Function


'--------------------------------------------------------------------------------
' Procedure  :       ColorARGB
' Description:       Creates an Alpha plus RGB colour value to use in GDI+
' Parameters :       Alpha (Byte)
'                    Red (Byte)
'                    Green (Byte)
'                    Blue (Byte)
'--------------------------------------------------------------------------------
Public Function ColorARGB(ByVal Alpha As Byte, _
                          ByVal Red As Byte, _
                          ByVal Green As Byte, _
                          ByVal Blue As Byte) As Long

    Dim bytestruct As COLORBYTES

    Dim result     As COLORLONG
   
    With bytestruct
        .AlphaByte = Alpha
        .RedByte = Red
        .GreenByte = Green
        .BlueByte = Blue
    End With
   
    LSet result = bytestruct
    ColorARGB = result.longval
End Function

