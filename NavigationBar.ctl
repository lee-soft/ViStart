VERSION 5.00
Begin VB.UserControl NavigationBar 
   BackColor       =   &H00F0F0F0&
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   Begin ViStart.MenuItem Item 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Caption         =   "Dummy Item"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line endLine 
      BorderColor     =   &H00CCCCCC&
      X1              =   216
      X2              =   216
      Y1              =   304
      Y2              =   0
   End
End
Attribute VB_Name = "NavigationBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private m_selectedIndex As Long

Private Const ROLLOVER_COLOUR As Long = &HE75314
Private Const NORMAL_COLOUR As Long = vbBlack
Private Const SELECTED_COLOUR As Long = &HC28854
Private Const SELECTED_FONT_COLOUR As Long = vbWhite

Public Event onClick(theIndex As Long)

Public Property Let SelectedIndex(newIndex As Long)
    Item_Click CInt(newIndex)
End Property

Sub ClearList()
    If Item.count = 1 Then Exit Sub
    
Dim labelIndex As Long

    For labelIndex = 1 To Item.UBound
        Unload Item(labelIndex)
    Next
End Sub

Function GetItemText(itemIndex As Long) As String

    If itemIndex >= Item.LBound And itemIndex <= Item.UBound Then
        GetItemText = Item(itemIndex).Caption
    End If

End Function

Function NavigateToItem(ByVal szItemCaption As String)

Dim itemIndex As Long
    szItemCaption = LCase$(szItemCaption)

    For itemIndex = Item.LBound To Item.UBound
        If LCase$(Item(itemIndex).Caption) = szItemCaption Then
            Item_Click CInt(itemIndex)
            Exit For
        End If
    Next

End Function

Sub AddItem(szText As String, Optional szTag As String)

Dim previousItemIndex As Long
Dim nextItem As MenuItem

    previousItemIndex = Item.UBound
    Load Item(Item.count)
    
    With Item(Item.UBound)
        .Top = Item(previousItemIndex).Top + CalculateItemGap
        .Left = Item(previousItemIndex).Left
        
        .Height = Item(previousItemIndex).Height
        .Width = Item(previousItemIndex).Width
        '.Tag = szTag
        
        .Caption = szText
        .Visible = True
    End With
End Sub

Private Function CalculateItemGap() As Long
    CalculateItemGap = Item(0).Height + 1
End Function

Private Sub Item_Click(index As Integer)
    Item(m_selectedIndex).BackColor = UserControl.BackColor
    Item(m_selectedIndex).ForeColor = NORMAL_COLOUR
    
    m_selectedIndex = index
    Item(m_selectedIndex).BackColor = SELECTED_COLOUR
    Item(m_selectedIndex).ForeColor = SELECTED_FONT_COLOUR
    
    RaiseEvent onClick(index - 1)
End Sub

Private Sub UserControl_Initialize()
    Item(0).Top = Item(0).Top - CalculateItemGap
    Item(0).Visible = False
End Sub

Private Sub UserControl_Resize()
    endLine.X1 = UserControl.ScaleWidth - 1
    endLine.X2 = endLine.X1
    
    endLine.Y1 = UserControl.ScaleHeight
    endLine.Y2 = 0
    
Dim labelIndex As Long

    For labelIndex = 0 To Item.UBound
        Item(labelIndex).Width = UserControl.ScaleWidth - Item(0).Left - 1
    Next
End Sub
