VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavigationPaneFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_parent As NavigationPaneOption

Public OpenAsMenu As Boolean

Public Property Get Shell() As String
    Shell = m_parent.Shell
End Property

Public Property Let Shell(newShell As String)
    m_parent.Shell = newShell
End Property

Public Property Let Visible(newVisible)
    m_parent.Visible = newVisible
End Property

Public Property Get Visible()
    Visible = m_parent.Visible
End Property

Public Property Get Position()
    Position = m_parent.Position
End Property

Public Property Let Position(newPosition)
    m_parent.Position = newPosition
End Property

Public Property Get Rollover()
    Rollover = m_parent.Rollover
End Property

Public Property Let Rollover(newRollover)
    m_parent.Rollover = newRollover
End Property

Public Property Get Caption()
    Caption = m_parent.Caption
End Property

Public Property Let Caption(newCaption)
    m_parent.Caption = newCaption
End Property

Private Sub Class_Initialize()
    Set m_parent = New NavigationPaneOption
End Sub

Private Property Let NavigationPaneOption_Caption(ByVal RHS As String)
    m_parent.Caption = RHS
End Property

Private Property Get NavigationPaneOption_Caption() As String
    NavigationPaneOption_Caption = m_parent.Caption
End Property

Private Property Let NavigationPaneOption_Position(ByVal RHS As Long)
    m_parent.Position = RHS
End Property

Private Property Get NavigationPaneOption_Position() As Long
    NavigationPaneOption_Position = m_parent.Position
End Property

Private Property Let NavigationPaneOption_Rollover(ByVal RHS As String)
    m_parent.Rollover = RHS
End Property

Private Property Get NavigationPaneOption_Rollover() As String
    NavigationPaneOption_Rollover = m_parent.Rollover
End Property

Private Property Let NavigationPaneOption_Shell(ByVal RHS As String)
    m_parent.Shell = RHS
End Property

Private Property Get NavigationPaneOption_Shell() As String
    NavigationPaneOption_Shell = m_parent.Shell
End Property

Private Property Let NavigationPaneOption_Visible(ByVal RHS As Boolean)
    m_parent.Visible = RHS
End Property

Private Property Get NavigationPaneOption_Visible() As Boolean
    NavigationPaneOption_Visible = m_parent.Visible
End Property
