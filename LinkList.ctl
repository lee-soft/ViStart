VERSION 5.00
Begin VB.UserControl LinkList 
   BackColor       =   &H00F0F0F0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MouseIcon       =   "LinkList.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   960
      Top             =   1080
   End
   Begin VB.Label Items 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Item 1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "LinkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_lastIndex As Long
Private m_currentPosition As RECT
Private m_selectedIndex As Long

Private Const ROLLOVER_COLOUR As Long = &HE75314
Private Const NORMAL_COLOUR As Long = vbBlack
Private Const SELECTED_COLOUR As Long = &H4AA182
Private Const SELECTED_FONT_COLOUR As Long = vbWhite

Public Event onClick(newIndex As Long)

Public Function SelectedItem() As Label
    If m_lastIndex > 0 And m_lastIndex <= Items.UBound Then
        Set SelectedItem = Items(m_lastIndex)
    End If
End Function

Sub ClearList()
    If Items.count = 1 Then Exit Sub
    
Dim labelIndex As Long

    For labelIndex = 1 To Items.UBound
        Unload Items(labelIndex)
    Next
End Sub

Sub AddItem(szText As String, Optional szTag As String)

Dim previousItem As Label
Dim nextItem As Label

    Set previousItem = Items(Items.UBound)
    Load Items(Items.count)
    Set nextItem = Items(Items.UBound)
    
    With nextItem
        .Top = previousItem.Top + previousItem.Height + 1
        .Left = previousItem.Left
        .Height = previousItem.Height
        .Width = UserControl.ScaleWidth - 10
        .Tag = szTag
        
        .Caption = szText
        .Visible = True
    End With
End Sub

Private Sub RolloutCurrent()
    If m_selectedIndex = m_lastIndex Then
        Items(m_lastIndex).ForeColor = SELECTED_FONT_COLOUR
    Else
        Items(m_lastIndex).ForeColor = NORMAL_COLOUR
    End If
    
    Items(m_lastIndex).FontUnderline = False
    
    m_lastIndex = 0
End Sub

Private Sub Items_Click(index As Integer)
    Items(m_selectedIndex).BackColor = UserControl.BackColor
    Items(m_selectedIndex).ForeColor = NORMAL_COLOUR
    
    m_selectedIndex = index
    Items(m_selectedIndex).BackColor = SELECTED_COLOUR
    Items(m_selectedIndex).ForeColor = SELECTED_FONT_COLOUR
    
    RaiseEvent onClick(index - 1)
End Sub

Private Sub Items_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_lastIndex = CLng(index) Then
        Exit Sub
    End If
    
    If Timer1.Enabled = False Then
        Timer1.Enabled = True
    End If
    
    RolloutCurrent
    m_lastIndex = CLng(index)
    
    If m_lastIndex <> m_selectedIndex Then
        Items(m_lastIndex).FontUnderline = True
        Items(m_lastIndex).ForeColor = ROLLOVER_COLOUR
    End If
End Sub

Private Sub Timer1_Timer()
    'Check cursor is inside window

Dim cursorPosition As POINTL
    cursorPosition = GetCursorPosition()

    GetWindowRect UserControl.hWnd, m_currentPosition
    
    If Not (cursorPosition.X > m_currentPosition.Left And cursorPosition.X < m_currentPosition.Right And _
         cursorPosition.Y > m_currentPosition.Top And cursorPosition.Y < m_currentPosition.Bottom) Then
        
        RolloutCurrent
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    Items(0).Top = -Items(0).Height
    Items(0).Visible = False
End Sub
