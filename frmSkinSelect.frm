VERSION 5.00
Begin VB.Form frmSkinSelect 
   Caption         =   "Select skin"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      Height          =   735
      Left            =   0
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   2
      Top             =   5640
      Width           =   8055
      Begin VB.CommandButton cmdClose 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Get More..."
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
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.ListBox lstSkins 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8055
   End
   Begin VB.Label lbltitle 
      Caption         =   "Select a new skin from the list and then select an option below."
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmSkinSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_skinDir As String

Public Event onChangeSkin(szNewSkin As String)

Sub ValidateSkin(ByVal szNewSkin As String)

    If FileCheck(m_skinDir & szNewSkin & "\") Then
        If Settings.CurrentSkin = vbNullString Or FileCheck(m_skinDir & Settings.CurrentSkin & "\") = False Then
            Settings.CurrentSkin = szNewSkin
        Else
            RaiseEvent onChangeSkin(szNewSkin)
        End If
        
        lbltitle.Caption = szNewSkin & " is the current skin."
    Else
        lbltitle.Caption = "This skin is broken!"
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If lstSkins.ListCount = 1 Then
        MsgBox "We need at least one skin", vbCritical
        Exit Sub
    End If

    If MsgBox("This operation cannot be undone." & vbCrLf & "Are you sure want to delete this skin?", vbYesNo Or vbExclamation) = vbNo Then
        Exit Sub
    End If

Dim szSkinName As String
Dim szSkinIndex As Long
    
    szSkinName = lstSkins.Text
    szSkinIndex = lstSkins.listIndex
    
    If szSkinIndex > 0 Then
        lstSkins.Selected(0) = True
    Else
        lstSkins.Selected(1) = True
    End If
    
    On Error Resume Next
    
    FSO.DeleteFolder m_skinDir & szSkinName, True
    If FSO.FolderExists(m_skinDir & szSkinName) = False Then
        lstSkins.RemoveItem szSkinIndex
        lbltitle.Caption = szSkinName & " was just deleted."
    Else
        If Err Then MsgBox Err.Description, vbCritical
        ListSkins
    End If
End Sub

Private Sub cmdSelect_Click()
    AppLauncherHelper.ShellEx "http://lee-soft.com/skins/"
End Sub

Private Sub Form_Initialize()

    Me.Font.Name = OptionsHelper.PrimaryFont
    
    lbltitle.Font.Name = OptionsHelper.PrimaryFont
    lstSkins.Font.Name = Me.Font.Name

End Sub

Sub ListSkins()
    On Error GoTo Handler

    lstSkins.Clear

Dim thisSubFolder As Scripting.Folder
Dim thisFolder As Scripting.Folder

    If FSO.FolderExists(m_skinDir) = False Then
        MsgBox "No skins available!", vbCritical
        End
    End If

    Set thisFolder = FSO.GetFolder(m_skinDir)
    
    For Each thisSubFolder In thisFolder.SubFolders
        If FileCheck(m_skinDir & thisSubFolder.Name & "\") Then
            lstSkins.AddItem thisSubFolder.Name
            
            If LCase$(thisSubFolder.Name) = LCase$(Settings.CurrentSkin) Then
                lstSkins.Selected(lstSkins.ListCount - 1) = True
            End If
        End If
    Next
    
    If lstSkins.ListCount = 0 Then
        MsgBox "No skins available!", vbCritical
        End
    End If

    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    m_skinDir = sCon_AppDataPath & "_skins\"
    
    ListSkins
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstSkins.Move 0, lstSkins.Top, Me.ScaleWidth, Me.ScaleHeight - 84
    picOptions.Move 0, Me.ScaleHeight - 50, Me.ScaleWidth
End Sub

Private Sub lstSkins_Click()
    ValidateSkin lstSkins.Text
End Sub

Private Sub picOptions_Resize()
    cmdClose.Move picOptions.ScaleWidth - 102
    'cmdInstall.Move picOptions.ScaleWidth - 205
End Sub
