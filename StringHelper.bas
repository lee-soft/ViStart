Attribute VB_Name = "StringHelper"
Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        m_logger = LogManager.GetLogger("StringHelper")
    End If
    
    Set Logger = m_logger
End Property


Function JustExe(ByVal szPath As String) As String

Dim extensionStart As Long
Dim extensionEnd As Long

    extensionStart = InStrRev(szPath, ".")

    If extensionStart < 0 Then extensionStart = 1
    
    If extensionStart < Len(szPath) Then
        extensionEnd = InStr(extensionStart, szPath, " ")
        If extensionEnd = 0 Then
            extensionEnd = Len(szPath)
        End If
        
        If extensionStart > 0 Then
            szPath = Left$(szPath, extensionEnd)
        End If
    End If

    'Kill Quotes
    szPath = Replace(szPath, """", vbNullString)
    szPath = Replace(szPath, "'", vbNullString)
    
    JustExe = szPath
End Function

Private Function Unserialize_INode(szData As String) As INode

Dim returnNode As New INode
    Set returnNode.Icon = New ViIcon

Dim sP() As String

    sP = Split(szData, "&")

    With returnNode
        .Caption = URLDecode(sP(0))
        .IconPosition = URLDecode(sP(1))
        .Tag = URLDecode(sP(2))
        .Icon.IconPath = URLDecode(sP(3))
    End With
    
    Set Unserialize_INode = returnNode
End Function

Private Function Serialize_INode(theINode As INode) As String
    
Dim szReturn As String
Dim szIcon As String

    If Not theINode.Icon Is Nothing Then
        szIcon = theINode.Icon.IconPath
    End If
 
    With theINode
        szReturn = szReturn & _
            "INode" & "?" & _
            URLEncode(.Caption) & "&" & _
            URLEncode(.IconPosition) & "&" & _
            URLEncode(.Tag) & "&" & _
            URLEncode(szIcon)
    End With
    
    Serialize_INode = szReturn
End Function

Function UnSerialize(szData As String)

Dim theType As String
Dim sP() As String

    sP = Split(szData, "?")
    If IsArrayInitialized(sP) Then
        If UBound(sP) = 1 Then
            
            Select Case UCase$(sP(0))
            
            Case "INODE"
                Set UnSerialize = Unserialize_INode(sP(1))
                
            Case Else
                MsgBox "'" & szData & "' cannot be unserialized!", vbCritical
            
            End Select
            
        End If
    End If

End Function

Function Serialize(theObject As Object) As String

    Select Case TypeName(theObject)
    
    Case "INode"
      Serialize = Serialize_INode(theObject)
      
    End Select

End Function

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function ExtOrNot(szFileName As String) As String

Dim periodPosition As Long
    periodPosition = InStrRev(szFileName, ".")
    
    If periodPosition > 0 Then
        ExtOrNot = Left$(szFileName, periodPosition - 1)
    Else
        ExtOrNot = szFileName
    End If

End Function

Public Function StrEnd(sData As String, sDelim As String, Optional iOffset As Integer = 1)

    If InStr(sData, sDelim) = 0 Then
        'Delim not present
    
        StrEnd = sData
        Exit Function
    End If

Dim iLen As Integer, iDLen As Integer

    iLen = Len(sData) + 1
    iDLen = Len(sDelim)

    If iLen = 1 Or iDLen = 0 Then
        StrEnd = False
        Exit Function
    End If

    While Mid$(sData, iLen, iDLen) <> sDelim And iLen > 1
        iLen = iLen - 1
    Wend

    If iLen = 0 Then
        StrEnd = False
        Exit Function
    End If
    
    StrEnd = Mid$(sData, iLen + iOffset)

End Function

Public Function GetDWord(Word As String, Optional Little As Boolean = False) As Double

    If LenB(Word) = 4 Then
        'Reverse if this is little
        If Little Then
            Dim i As Long, SStr As String
            For i = Len(Word) To 1 Step -1
                SStr = SStr & MidB$(Word, i, 1)
            Next i
            Word = SStr
        End If
        
        'Grab hex values
        Dim J As Long, HStr As String, H As String
        For J = 1 To Len(Word)
            H = Hex$(AscB(MidB$(Word, J, 1)))
            If Len(H) = 1 Then H = "0" & H
            HStr = HStr & H
        Next J
        
        'Cut off padding (null characters)
        Do
            HStr = Left$(HStr, Len(HStr) - 2)
        Loop While Right$(HStr, 2) = "00"

        'Return the value
        GetDWord = Val("&H" & IIf(HStr <> "", HStr, "00") & "&")
    Else
        'No correct Word supplied
        GetDWord = 0
    End If
End Function

Function StrToHex(ByRef str)
    Dim Length
    Dim Max
    Dim strHex
    Max = Len(str)
    For Length = 1 To Max
        strHex = strHex & Right$("0" & Hex$(Asc(Mid$(str, Length, 1))), 2)
    Next
    StrToHex = strHex
End Function

Function GetStringByPosition(ByRef sSource As String, ByVal lngPos As Long) As String
    
Dim sNewString As String
    
    If lngPos > 0 Then
        sNewString = Mid$(sSource, 1, lngPos - 1)
        
        GetStringByPosition = sNewString
        sSource = MidB$(sSource, lngPos)
    End If
    
End Function

Public Function ExtractBytes(ByRef strUniSource, lngBytes As Long) As String
    
Dim strBuffer As String

    strBuffer = MidB$(strUniSource, 1, lngBytes)
    strUniSource = MidB$(strUniSource, lngBytes + 1)
    
    ExtractBytes = strBuffer
    
End Function

Function GetStringByString(ByRef sSource As String, ByVal sDelim As String) As String
    
Dim lngPos As Long
Dim sNewString As String
    
    lngPos = InStr(sSource, sDelim)
    
    If lngPos > 0 Then
        sNewString = Mid$(sSource, 1, lngPos - 1)
        
        GetStringByString = sNewString
        sSource = Mid$(sSource, lngPos + Len(sDelim))
    End If
    
End Function

Public Function CBol(ByRef vData) As Boolean
    If vData = "1" Or vData = "True" Then
        CBol = True
    End If
End Function


Public Function HEXCOL2RGB(ByVal HexColor As String) As String

    'The input at this point could be HexColor = "#00FF1F"
Dim Red As String
Dim Green As String
Dim Blue As String

    On Error GoTo Handler

HexColor = Replace(HexColor, "#", "")
    'Here HexColor = "00FF1F"

Red = Val("&H" & Mid$(HexColor, 1, 2))
    'The red value is now the long version of "00"

Green = Val("&H" & Mid$(HexColor, 3, 2))
    'The red value is now the long version of "FF"

Blue = Val("&H" & Mid$(HexColor, 5, 2))
    'The red value is now the long version of "1F"


HEXCOL2RGB = RGB(Red, Green, Blue)
    'The output is an RGB value
    
    Exit Function
Handler:
    HEXCOL2RGB = vbWhite

End Function

Public Function isChecked(ByRef bBol As Boolean) As Long
    If bBol Then
        isChecked = 1
    Else
        isChecked = 0
    End If
End Function

Public Function HexToString(ByVal HexToStr As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim i         As Long
    For i = 1 To Len(HexToStr) Step 2
        strTemp = ChrB$(Val("&H" & Mid$(HexToStr, i, 2)))
        strReturn = strReturn & strTemp
    Next i
    HexToString = strReturn
End Function

Public Function ExistInStringArray(ByRef theArray() As String, ByVal theDelimiter As String) As Boolean
Dim arrayIndex As Long
    theDelimiter = UCase$(theDelimiter)
    
    If Not isset(theArray) Then Exit Function
    For arrayIndex = LBound(theArray) To UBound(theArray)
        If UCase$(theArray(arrayIndex)) = theDelimiter Then
            ExistInStringArray = True
            Exit For
        End If
    Next
    
End Function
