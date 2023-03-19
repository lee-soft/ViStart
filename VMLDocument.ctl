VERSION 5.00
Begin VB.UserControl VMLDocument 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   Begin VB.Timer timRefresh 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   630
      Top             =   1050
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "<PLACE HOLDER>"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   20
      MouseIcon       =   "VMLDocument.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Images 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2565
   End
End
Attribute VB_Name = "VMLDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: VMLDocument
' Kind: UserControl
' Purpose: Provide basic HTML rendering using VB6 components
' Author: Lee Chantrey
' Date: 17/03/2023
' ------------------------------------------------------

Option Explicit

Private Enum FontStyle
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleRegular = 4
    FontStyleStrikeout = 5
    FontStyleUnderline = 6
End Enum

Private m_ElementY As Long
Private m_ElementX As Long

Private m_sXML As String
Private m_Margin As RECT
Private m_lineIsDirty As Boolean

Private m_lAvailbleWidth As Long

Public Event LinkClicked(ByVal tag As String)



Private Sub ClearElements()
    On Error GoTo Handler
    
    While Labels.count > 1
        Unload Labels(Labels.UBound)
    Wend
    
    While Images.count > 1
        Unload Images(Images.UBound)
    Wend
Handler:
    Debug.Print Err.Description
End Sub

Public Property Get XML() As String
    XML = m_sXML
End Property

Public Property Let XML(ByVal sNewValue As String)
    If m_sXML = sNewValue Then Exit Property
    
    m_sXML = sNewValue
    RenderXML
End Property

Private Sub RenderXML()

    ClearElements

Dim xmlDoc As DOMDocument
Dim skinInfoXML As IXMLDOMElement
Dim thisChild As IXMLDOMElement
Dim thisObject As Object

    Set xmlDoc = New DOMDocument
    
    If xmlDoc.loadXML(m_sXML) = False Then
        Exit Sub
    End If
 
    m_ElementY = m_Margin.Top
    m_ElementX = m_Margin.Left
    
    Set skinInfoXML = xmlDoc.firstChild
    
    If Not IsNull(skinInfoXML.getAttribute("margin")) Then
        m_Margin.Left = skinInfoXML.getAttribute("margin")
        m_Margin.Right = skinInfoXML.getAttribute("margin")
        m_Margin.Top = skinInfoXML.getAttribute("margin")

    End If
    
    For Each thisObject In skinInfoXML.childNodes
        ParseChildXML thisObject
    Next
    
End Sub

Private Sub ParseChildXML(thisObject As Object, Optional parentObject As IXMLDOMElement = Nothing)

Dim thisChild As IXMLDOMElement

    Select Case TypeName(thisObject)
    
    Case "IXMLDOMElement"
        Set thisChild = thisObject

        Select Case LCase$(thisChild.tagName)
        
        Case "a"
            ParseHref thisChild
        
        Case "p"
            ParseParagraph thisChild
            
        Case "img"
            ParseImage thisChild, parentObject
            
        Case "h1"
            ParseHeader thisChild
        
        End Select
    
    Case "IXMLDOMText"
        ParseText thisObject, parentObject
        
    End Select

End Sub

Private Sub ParseImage(ByRef theXMLImage As IXMLDOMElement, Optional ByRef theParent As IXMLDOMElement = Nothing)

Dim objXML As MSXML2.DOMDocument
Dim objNode As MSXML2.IXMLDOMElement

Set objXML = New MSXML2.DOMDocument
Set objNode = objXML.createElement("b64")
objNode.dataType = "bin.base64"

Dim theAlignment As AlignmentConstants

    If Not theParent Is Nothing Then
    
        Dim theHref As String

        If Not IsNull(theParent.getAttribute("href")) Then
            theHref = theParent.getAttribute("href")
        End If
        
        If Not IsNull(theParent.getAttribute("align")) Then
            Select Case LCase(theParent.getAttribute("align"))
            
            Case "left"
                theAlignment = vbLeftJustify
            
            Case "right"
                theAlignment = vbRightJustify
                
            Case "center"
                theAlignment = vbCenter
                
            End Select
        End If
    
    End If

    If Not IsNull(theXMLImage.getAttribute("src")) Then
        ' Set the text value of the element to the Base64 string
        objNode.text = theXMLImage.getAttribute("src")
    End If
    
    If Not IsNull(theXMLImage.getAttribute("align")) Then
        Select Case LCase(theXMLImage.getAttribute("align"))
        
        Case "left"
            theAlignment = vbLeftJustify
        
        Case "right"
            theAlignment = vbRightJustify
            
        Case "center"
            theAlignment = vbCenter
            
        End Select
    End If

Dim rawBinary() As Byte
    rawBinary = objNode.nodeTypedValue

Dim sourceImage As Image
    Set sourceImage = AddImage(PictureFromBits(rawBinary), theAlignment)
    
    If theHref <> vbNullString Then
        With sourceImage
            .tag = theHref
            .MousePointer = 99
        End With
    End If
End Sub

Private Function AddImage(picture As IPicture, theAlignment As AlignmentConstants) As Image

Dim sourceImage As Image
Dim imageLeft As Long

Dim fullTextWidth As Long: fullTextWidth = AvailableWidth

    Load Images(Images.count)
    Set sourceImage = Images(Images.UBound)

    With sourceImage
        .MouseIcon = Labels(0).MouseIcon
        .picture = picture
        .Stretch = False
        .Top = m_ElementY
  
        m_ElementY = m_ElementY + .Height
    End With
    
    If theAlignment = vbRightJustify Then
        imageLeft = (m_Margin.Left + fullTextWidth) - sourceImage.Width
    ElseIf theAlignment = vbCenter Then
        imageLeft = ((m_Margin.Left + (fullTextWidth / 2)) - sourceImage.Width / 2)
    ElseIf theAlignment = vbLeftJustify Then
        imageLeft = m_Margin.Left
    End If

    sourceImage.Left = imageLeft
    sourceImage.Visible = True

    Set AddImage = sourceImage
End Function

Private Sub ParseHref(ByRef theText As IXMLDOMElement)
    On Error GoTo Handler
    ParseChildXML theText.firstChild, theText

    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub ParseText(ByRef theText As IXMLDOMText, Optional ByRef theParent As IXMLDOMElement = Nothing)
    
    If theParent Is Nothing Then
        AddText theText.text, FontStyleRegular, vbLeftJustify
        Labels(0).Caption = theText.text
        m_ElementX = m_ElementX + Labels(0).Width + UserControl.TextWidth(" ")
        m_lineIsDirty = True
        Exit Sub
    End If
    
    If (theParent.firstChild.text <> vbNullString) Then
        
    Dim theCaption As String
    Dim theHref As String
    Dim theAlignment As AlignmentConstants
    Dim sourceLabel As Label
    
        theAlignment = vbLeftJustify
    
        If Not IsNull(theParent.getAttribute("align")) Then
            Select Case LCase(theParent.getAttribute("align"))
            
            Case "left"
                theAlignment = vbLeftJustify
            
            Case "right"
                theAlignment = vbRightJustify
                
            Case "center"
                theAlignment = vbCenter
                
            End Select
        End If
            
        
        If theParent.tagName = "h1" Then
        
            theCaption = theText.text
            Set sourceLabel = AddText(theCaption, FontStyleRegular, theAlignment)
            
            Labels(0).AutoSize = True
            Labels(0).fontSize = 15
            Labels(0).Caption = theCaption
            Labels(0).AutoSize = True
        
            With sourceLabel
                .tag = theHref
                .fontSize = 15
                .Height = Labels(0).Height
            End With
            
            Labels(0).fontSize = 9
            
            m_ElementY = m_ElementY + 25
        
        Else
            theCaption = theText.text
        
            If Not IsNull(theParent.getAttribute("href")) Then
                theHref = theParent.getAttribute("href")
            End If

            Set sourceLabel = AddText(theCaption, FontStyleRegular, theAlignment)
        
            With sourceLabel
                .tag = theHref
                .FontUnderline = True
                .ForeColor = vbBlue
                .MousePointer = 99
                .ZOrder 0
            End With
            
            Labels(0).Caption = theText.text
            m_ElementX = m_ElementX + Labels(0).Width + UserControl.TextWidth(" ")
            m_lineIsDirty = True
            
        End If
    Else
        ParseChildXML theText, theText
    End If
End Sub

Private Sub ParseHeader(ByRef theText As IXMLDOMElement)
    On Error GoTo Handler
    ParseChildXML theText.firstChild, theText

    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub ParseParagraph(ByRef theText As IXMLDOMElement)
    'CarriageReturnLineFeed
    CarriageReturnLineFeed
    
Dim theCaption As String
Dim theFontStyleText As String
Dim realPosition_X As Long

Dim theLabelIndex As Long
Dim theFontStyle As FontStyle
Dim theAlignment As AlignmentConstants

    theAlignment = vbLeftJustify

    If Not IsNull(theText.text) Then
        theCaption = theText.text
    End If
    
    If Not IsNull(theText.getAttribute("align")) Then
        Select Case LCase(theText.getAttribute("align"))
        
        Case "left"
            theAlignment = vbLeftJustify
        
        Case "right"
            theAlignment = vbRightJustify
            
        Case "center"
            theAlignment = vbCenter
            
        End Select
    End If
    
    If Not IsNull(theText.getAttribute("style")) Then
        theFontStyleText = theText.getAttribute("style")
    End If
    
    theFontStyle = FontStyleRegular
    
    Select Case LCase$(theFontStyleText)
    
    Case "bold"
        theFontStyle = FontStyleBold
        
    Case "italic"
        theFontStyle = FontStyleItalic
        
    Case "bold|italic"
        theFontStyle = FontStyleBoldItalic
        
    Case "underline"
        theFontStyle = FontStyleUnderline
        
    Case "strikeout"
        theFontStyle = FontStyleStrikeout
    
    End Select
    
Dim sourceLabel As Label
    
    Set sourceLabel = AddText(theCaption, theFontStyle, theAlignment)
    
    m_ElementY = m_ElementY + sourceLabel.Height
        
    'AddText "", FontStyleRegular, vbLeftJustify
    
    CarriageReturnLineFeed
End Sub

Public Property Get AvailableWidth() As Long
    AvailableWidth = (UserControl.ScaleWidth - (m_Margin.Left + m_Margin.Right))
End Property

Private Function CalculateNumberOfLines(ByVal strText As String) As Long

Dim intLineCount As Integer
Dim intTextWidth As Integer
Dim intTextHeight As Integer
Dim strWords() As String
Dim intLineWidth As Integer
Dim i As Integer

Dim fullTextWidth As Long: fullTextWidth = AvailableWidth

' Split the input string into words
strWords = Split(strText, " ")

' Initialize the line width and line count
intLineWidth = 0
intLineCount = 1

Labels(0).Caption = ""

' Loop through each word and calculate the total line width
For i = 0 To UBound(strWords)
    ' Calculate the width of the current word
    ' Add the word width to the line width
    
    Labels(0).AutoSize = True
    Labels(0).Caption = Labels(0).Caption + strWords(i) + " "
    intLineWidth = Labels(0).Width
    
    ' If the line width exceeds the available width, start a new line
    If intLineWidth > fullTextWidth Then
        If i - 1 > UBound(strWords) Then i = i - 1
        
        Labels(0).Caption = ""
        intLineCount = intLineCount + 1
    End If
Next i

CalculateNumberOfLines = intLineCount

End Function

Private Function CarriageReturnLineFeed()

Dim doubleDown As Boolean

    If m_lineIsDirty Then
        doubleDown = True
    End If
    
    Labels(0).AutoSize = True
    Labels(0).Caption = "A"
    m_ElementY = m_ElementY + Labels(0).Height
    m_ElementX = m_Margin.Left
    m_lineIsDirty = False
    
    If doubleDown Then CarriageReturnLineFeed
End Function

Private Function AddText(ByVal theCaption As String, Optional ByVal theFontStyle As FontStyle, Optional TextAlignment As AlignmentConstants = vbLeftJustify) As Label
    On Error GoTo Handler

Dim newHeight As Long
Dim lineHeight As Long

    Labels(0).AutoSize = True
    Labels(0).Caption = theCaption
    lineHeight = Labels(0).Height
    
    'TODO: add intercepter in the form of an event to allow consumer to overide captions
    theCaption = Replace(theCaption, "%skin%", Settings.CurrentSkin)

    Load Labels(Labels.count)

    With Labels(Labels.UBound)
        .Left = m_ElementX
        
        .ForeColor = vbBlack
        
        .AutoSize = False
        .Height = CalculateNumberOfLines(theCaption) * lineHeight

        
        If theFontStyle = FontStyleBold Then
            .FontBold = True
        ElseIf theFontStyle = FontStyleBoldItalic Then
            .FontBold = True
            .FontItalic = True
        ElseIf theFontStyle = FontStyleItalic Then
            .FontItalic = True
        ElseIf theFontStyle = FontStyleStrikeout Then
            .FontStrikethru = True
        ElseIf theFontStyle = FontStyleUnderline Then
            .FontUnderline = True
        End If
        
        .Caption = theCaption
        .MousePointer = 0

        .Visible = True
        .Width = AvailableWidth

        .Top = m_ElementY
        .Caption = .Caption
        .Alignment = TextAlignment
        
        '.BackStyle = 1
        '.BackColor = vbRed
    End With
    
    Set AddText = Labels(Labels.UBound)
    
    Exit Function
Handler:
    MsgBox Err.Description, vbCritical
End Function

Private Sub Images_Click(Index As Integer)
    RaiseEvent LinkClicked(Images(Index).tag)
End Sub

Private Sub Labels_Click(Index As Integer)
    RaiseEvent LinkClicked(Labels(Index).tag)
End Sub

Private Sub timRefresh_Timer()
    timRefresh.Enabled = False
    RenderXML

End Sub

Private Sub UserControl_Initialize()
    m_Margin.Left = 40
    m_Margin.Right = 40
    m_Margin.Top = 20
End Sub

Private Sub UserControl_Resize()
    timRefresh.Enabled = True
End Sub
