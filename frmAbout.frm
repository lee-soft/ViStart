VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ViStart"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timSplashMin 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3720
   End
   Begin VB.Timer timUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   180
      Top             =   3690
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4380
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   4380
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      Picture         =   "frmAbout.frx":74F2
      ScaleHeight     =   705
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This version of ViStart has been rewritten. Some portions of this software are based on the orignal ViStart."
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lee Matthew Chantrey"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lchantrey@gmail.com"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1350
      TabIndex        =   10
      Top             =   4560
      Width           =   2325
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.lee-soft.com"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1760
      MouseIcon       =   "frmAbout.frx":8FA4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4320
      Width           =   1545
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Build 2340"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblBottom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I hope you enjoy using ViStart"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3900
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":92AE
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lee-Soft is in no way associated with Microsoft or Windows."
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ViStart 8.1"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/vistart/"
End Sub

Private Sub Form_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub Form_Load()
    Label3.Caption = "Lee Matthew Chantrey"
    Label2.Caption = "(build " & App.Revision & ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If timUpdate.Enabled Then
        Cancel = 1
    End If
End Sub

Private Sub Label1_Click()
    AppLauncherHelper.ShellEx "http://nightly.lee-soft.com/"
End Sub

Private Sub Label2_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub Label3_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub Label4_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub Label5_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub Label6_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub lblBottom_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub lblLink_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub Picture1_Click()
    AppLauncherHelper.ShellEx "http://www.lee-soft.com/"
End Sub

Private Sub timSplashMin_Timer()
    timUpdate.Enabled = False
    Unload Me
End Sub

Private Sub timUpdate_Timer()
    lblEmail.Caption = lblEmail.Caption & "."
End Sub
