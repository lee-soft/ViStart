VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Trapper"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFunction 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "&Terminate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtReason 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox txtModule 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Function"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4320
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label3 
      Caption         =   "Reason"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Module/Class Name"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ShowError(sModule As String, sFunction As String, sReason As String)

    txtModule.Text = sModule
    txtFunction.Text = sFunction
    txtReason.Text = sReason
    
    'cmdIgnore_Click
    Me.Show

End Sub

Private Sub cmdIgnore_Click()
    Unload Me

End Sub

Private Sub cmdTerminate_Click()
    ExitApplication
End Sub

