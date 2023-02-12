VERSION 5.00
Begin VB.UserControl SettingsOption 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   Begin VB.ComboBox cmbOptions 
      Height          =   405
      ItemData        =   "LayoutOption.ctx":0000
      Left            =   1560
      List            =   "LayoutOption.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "SettingsOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function DrawTextW Lib "user32.dll" _
    (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private m_Text As String
Private m_labelPosition As RECT
Private m_Source As Object

Public Event onChanged()

Private Const DONT_SHOW_ITEM As Long = 0
Private Const DISPLAY_AS_LINK As Long = 1
Private Const DISPLAY_AS_MENU As Long = 2


Public Function ComitChanges()
    If m_Source Is Nothing Then Exit Function

    m_Source.Visible = Me.OptionVisible
    
    If TypeName(m_Source) = "NavigationPaneFolder" Then
        m_Source.OpenAsMenu = Me.OpenAsMenu
    End If
End Function

Public Property Let Source(newSource As Object)
    Set m_Source = newSource
    
    Me.Caption = VarScan(newSource.Caption)
    
    If newSource.Visible = False Then
        cmbOptions.listIndex = DONT_SHOW_ITEM
        Exit Property
    End If
    
    If TypeName(m_Source) = "NavigationPaneFolder" Then
        If m_Source.OpenAsMenu Then
            cmbOptions.listIndex = DISPLAY_AS_MENU
        Else
            cmbOptions.listIndex = DISPLAY_AS_LINK
        End If
    ElseIf TypeName(m_Source) = "NavigationPaneCustom" Then
        cmbOptions.RemoveItem 2
        
        If m_Source.Visible Then
            cmbOptions.listIndex = DISPLAY_AS_LINK
        Else
            cmbOptions.listIndex = 1
        End If
    End If
    

End Property

Public Property Get OptionVisible() As Boolean
    OptionVisible = IIf(cmbOptions.listIndex = DONT_SHOW_ITEM, False, True)
End Property

Public Property Get OpenAsMenu() As Boolean
    OpenAsMenu = IIf(cmbOptions.listIndex = DISPLAY_AS_MENU, True, False)
End Property

Public Property Let Caption(newCaption As String)
    m_Text = newCaption
    
    UserControl.Refresh
End Property

Private Sub cmbOptions_Change()
    ComitChanges
    RaiseEvent onChanged
End Sub

Private Sub cmbOptions_Click()
    ComitChanges
    RaiseEvent onChanged
End Sub

Private Sub UserControl_Initialize()
    cmbOptions.Clear
    
    cmbOptions.AddItem GetPublicString("strDontShowItem")
    cmbOptions.AddItem GetPublicString("strDisplayAsLink")
    cmbOptions.AddItem GetPublicString("strDisplayAsMenu")
    
    
    m_Text = "Test"
    m_labelPosition = CreateRect(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
End Sub

Private Sub UserControl_Paint()
    UserControl.Cls
    DrawTextW UserControl.hdc, StrPtr(m_Text), Len(m_Text), m_labelPosition, DT_LEFT Or DT_NOPREFIX Or DT_MODIFYSTRING
End Sub

Private Sub UserControl_Resize()
    m_labelPosition = CreateRect(5, 10, UserControl.ScaleWidth - cmbOptions.Width - 10, UserControl.ScaleHeight)
    cmbOptions.Move UserControl.ScaleWidth - cmbOptions.Width - 10

    UserControl.Refresh
End Sub
