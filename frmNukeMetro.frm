VERSION 5.00
Begin VB.Form frmNukeMetro 
   Caption         =   "Wait"
   ClientHeight    =   1350
   ClientLeft      =   3330
   ClientTop       =   10230
   ClientWidth     =   5040
   Icon            =   "frmNukeMetro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   5040
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nuking Metro...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmNukeMetro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_toolTip As ViToolTip

Private Sub Command1_Click()
    m_toolTip.SetToolTip "Test"
    m_toolTip.Show
End Sub

Private Sub Form_Load()
    Set m_toolTip = New ViToolTip
    m_toolTip.AttachWindow Me.hWnd
    m_toolTip.SetToolTip "Test"
    
    
End Sub

Private Sub Timer1_Timer()
    m_toolTip.Show
End Sub
