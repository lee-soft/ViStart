VERSION 5.00
Begin VB.Form frmVistaMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "Vista_FileMenu"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   ShowInTaskbar   =   0   'False
   Tag             =   "0"
   Begin VB.PictureBox imgSel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer timRollover 
      Interval        =   5
      Left            =   3480
      Top             =   3240
   End
   Begin VB.Image ImgSep 
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00B9B5B7&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblItems1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Top             =   -270
      Width           =   480
   End
   Begin VB.Line ln4 
      BorderColor     =   &H00FFFFFF&
      X1              =   168
      X2              =   168
      Y1              =   1
      Y2              =   64
   End
   Begin VB.Line ln3 
      BorderColor     =   &H00FFFFFF&
      X1              =   27
      X2              =   27
      Y1              =   1
      Y2              =   56
   End
   Begin VB.Line ln2 
      BorderColor     =   &H00E0E0DF&
      X1              =   26
      X2              =   26
      Y1              =   1
      Y2              =   80
   End
   Begin VB.Shape shpLeft 
      BorderColor     =   &H00EAEAEA&
      FillColor       =   &H00EAEAEA&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   30
      Top             =   30
      Width           =   360
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1
      X2              =   1
      Y1              =   1
      Y2              =   100
   End
End
Attribute VB_Name = "frmVistaMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawTextW Lib "user32.dll" _
    (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const C_HEIGHT As Integer = 22
Private Const C_SEP As Integer = 9

Private Const C_LEFTITEMSTART As Long = 30

Private Type VM_ITEM
    Caption As String
    Bold As Boolean
    Tag As String
    lpRect As RECT
End Type

Public iHideCount As Long

Private lblItems() As VM_ITEM
Private iCurIndex As Integer
Private iLblUbound As Long

Private meRect As RECT
Private recRollover As RECT

Private m_maxWidth As Long
Private m_maxHeight As Long
Private m_allowInActive As Boolean

Public Event onInActive()
Public Event onClick(theItemTag As String)

Implements IHookSink

Public Function CountItems() As Long
    CountItems = UBound(lblItems)
End Function

Public Function IsItemBold(ByVal itemIndex As Long) As Boolean
    IsItemBold = lblItems(itemIndex).Bold
End Function

Public Function GetItemCaption(ByVal itemIndex As Long) As String
    GetItemCaption = lblItems(itemIndex).Caption
End Function

Public Function GetItemExec(ByVal itemIndex As Long) As String
    GetItemExec = lblItems(itemIndex).Tag
End Function

Public Function SetItemCaption(ByVal itemIndex As Long, ByVal newCaption As String)
    lblItems(itemIndex).Caption = newCaption
End Function

Private Sub Command1_Click()
    PaintLabels
End Sub

Private Sub Form_Activate()

    GetWindowRect Me.hWnd, meRect
    Rollover iCurIndex
    'iHideCount = 0

End Sub

Sub SetControlProperties()

    Set imgSel.Picture = GetResourceBitmap(101)

    ReDim Preserve lblItems(0)
    lblItems(0).lpRect.Top = -18
    lblItems(0).lpRect.Left = 30
    
    Me.FontName = sVar_sFontName
    imgSel.FontName = sVar_sFontName
    
    With recRollover
        .Left = 30
        .Bottom = C_HEIGHT
        .Right = imgSel.ScaleWidth
        .Top = 4
    End With

End Sub

Private Sub Form_Initialize()
    SetControlProperties
    
    Me.Width = 1024 * Screen.TwipsPerPixelX
    HookWindow Me.hWnd, Me
    
End Sub

Private Sub Form_Resize()

    Me.AlignElements

End Sub

Sub PaintLabels()

Dim lngLabelIndex As Long
Dim lngSepIndex As Long
Dim recLength As RECT

    Me.Cls
    
    For lngLabelIndex = LBound(lblItems) To UBound(lblItems)
        FontName = OptionsHelper.PrimaryFont
        Me.FontBold = lblItems(lngLabelIndex).Bold
        
        'Debug.Print lngLabelIndex & " - " & lblItems(lngLabelIndex).Caption
        DrawText Me.hdc, lblItems(lngLabelIndex).Caption, lblItems(lngLabelIndex).lpRect, 0
    Next
    
    imgSel.Width = Me.ScaleWidth
    For lngSepIndex = 0 To ImgSep.UBound
        ImgSep(lngSepIndex).Width = Me.ScaleWidth - 1
    Next
    
    Rollover iCurIndex
End Sub

Sub AlignElements()
    On Error Resume Next

    shpBorder.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    ln1.Y2 = Me.ScaleHeight - 1
    
    shpLeft.Move 2, 1, 24, Me.ScaleHeight - 2
    
    ln1.X1 = 26
    ln1.X2 = 26
    
    ln2.X1 = 26
    ln2.X2 = 26
    
    ln2.Y2 = Me.ScaleHeight - 1
    ln3.Y2 = Me.ScaleHeight - 1
    
    ln3.X1 = 27
    ln3.X2 = 27
    
    ln4.X1 = Me.ScaleWidth - 2
    ln4.X2 = Me.ScaleWidth - 2
    ln4.Y2 = Me.ScaleHeight - 1

End Sub

Sub AddItem(sItem As String, Optional sShell As String, Optional Bold As Boolean = False)

Dim iNext As Integer
Dim newRect As RECT

    iLblUbound = iLblUbound + 1
    iNext = iLblUbound

    ReDim Preserve lblItems(iNext)

    lblItems(iNext).lpRect.Left = C_LEFTITEMSTART
    lblItems(iNext).Bold = Bold
    
    If sItem <> "" Then
        lblItems(iNext).lpRect.Top = m_maxHeight + 4
        lblItems(iNext).Tag = sShell
        
        m_maxHeight = m_maxHeight + C_HEIGHT
    Else
    
        lblItems(iNext).lpRect.Top = m_maxHeight - 8
        lblItems(iNext).Tag = "SEP"
        
        m_maxHeight = m_maxHeight + C_SEP
        AddSep
    End If

    lblItems(iNext).lpRect.Bottom = lblItems(iNext).lpRect.Top + C_HEIGHT
    lblItems(iNext).lpRect.Right = 999999
    lblItems(iNext).Caption = sItem
    
    'Debug.Print "Setting caption( " & iNext & " ):: " & lblItems(iNext).Caption

    'Calculate MaxWidth
    DrawText Me.hdc, sItem, newRect, DT_CALCRECT
    If newRect.Right > m_maxWidth Then
        m_maxWidth = newRect.Right
    End If

End Sub

Sub AddSep()

Dim imgNew As Image
Dim iNext As Integer

    iNext = ImgSep.count
    
    Load ImgSep(iNext)
    Set imgNew = ImgSep(iNext)
    
    imgNew.Picture = GetResourceBitmap(103)
    imgNew.Top = lblItems(iLblUbound).lpRect.Top + C_SEP
    imgNew.Visible = True
    imgNew.ZOrder 0
    
    shpBorder.ZOrder 0

End Sub

Private Function IHookSink_WindowProc(hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
    
    'form-specific handler
    Select Case uMsg
       
       Case WM_ACTIVATE
            'Debug.Print "WM_ACTIVATE!"
       
           If wParam = False Then
               'frmEvents.m_winVistaMenu_LostFocus
               If m_allowInActive Then RaiseEvent onInActive
           End If
    
       Case Else
           ' Just allow default processing for everything else.
           IHookSink_WindowProc = _
              CallOldWindowProcessor(hWnd, uMsg, wParam, lParam)
                                          
           Exit Function
       
    End Select
End Function

Private Sub imgSel_Click()

    lblItems_Click iCurIndex

End Sub

Private Sub lblItems_Click(index As Integer)
    
Dim sP() As String
Dim lngIndex As Long

    RaiseEvent onClick(lblItems(index).Tag)
End Sub

Private Sub timRollover_Timer()
    If Me.Visible = False Then
        Exit Sub
    End If

Dim iTop As Integer
Dim cPos As POINTL
Dim i As Long

    GetCursorPos cPos

    If cPos.X < meRect.Left Or _
        cPos.X > meRect.Right Or _
         cPos.Y < meRect.Top Or _
          cPos.Y > meRect.Bottom _
        Then
        
        iCurIndex = -1
        imgSel.Top = -(imgSel.Height * 2)
        Exit Sub
    End If
    
    For i = 0 To iLblUbound
        iTop = (Me.Top / Screen.TwipsPerPixelY) + lblItems(i).lpRect.Top - 4
    
        If cPos.Y < (iTop + 22) And _
                cPos.Y > (iTop) Then
                
                If lblItems(i).Tag <> "SEP" Then
                    If iCurIndex <> i Then
                        iCurIndex = i
                        
                        Rollover iCurIndex
                    End If
                End If
        End If
    Next

End Sub

Sub Rollover(index As Integer)
    If index < 0 Then
        imgSel.Top = -(imgSel.Height * 2)
        Exit Sub
    End If
    
    If index >= LBound(lblItems) And index <= UBound(lblItems) Then
    
        imgSel.Move 0, lblItems(index).lpRect.Top - 4
        imgSel.Cls
        imgSel.FontName = OptionsHelper.PrimaryFont
        imgSel.FontBold = lblItems(index).Bold
        
        Debug.Print lblItems(index).Caption
        DrawText imgSel.hdc, lblItems(index).Caption, recRollover, 0
    End If
End Sub

Sub Die()
    If Me.Visible Then
        Me.Tag = 0
        iHideCount = 0
        timRollover.Enabled = False

        Me.Hide
    End If
End Sub

Sub SetDimensions()

    PaintLabels

    Me.Width = (34 + (m_maxWidth) + 5) * Screen.TwipsPerPixelX
    Me.Height = m_maxHeight * Screen.TwipsPerPixelY

End Sub

Sub Resurrect(Optional WhereCursorIs As Boolean = False, Optional ownerForm As Form)

Dim lpCursor As POINTL

    Me.Tag = 0
    timRollover.Enabled = True
    
    PaintLabels
    EnsureIAmSeen
    
    Me.Width = (34 + (m_maxWidth) + 5) * Screen.TwipsPerPixelX
    Me.Height = m_maxHeight * Screen.TwipsPerPixelY
    
    If WhereCursorIs Then
        GetCursorPos lpCursor
        
        If (lpCursor.Y + Me.ScaleHeight) > (Screen.Height / Screen.TwipsPerPixelY) Then
            Debug.Print "HARTE!"
            lpCursor.Y = lpCursor.Y - Me.ScaleHeight
        End If
        
        'Fix window not re-appearing bug by adding -1, possible painting glitch
        Me.Move (lpCursor.X) * Screen.TwipsPerPixelX, (lpCursor.Y) * Screen.TwipsPerPixelY
    End If
    
    If Not ownerForm Is Nothing Then
        Me.Show vbModeless, ownerForm
    Else
        Me.Show
    End If
    
    TopMost Me.hWnd
    
    m_allowInActive = True
    
    PaintLabels
End Sub

Private Sub EnsureIAmSeen()

Dim lpCursor As POINTL

    GetWindowRect Me.hWnd, meRect
    If meRect.Right > (Screen.Width / Screen.TwipsPerPixelX) Then
        GetCursorPos lpCursor
        Me.Left = ((lpCursor.X - Me.ScaleWidth) * Screen.TwipsPerPixelX)
        
        'Reactive Form, at new position
        Form_Activate
    End If

End Sub

Private Function DrawText(hdc As Long, strStr As String, lpRect As RECT, wFormat As Long)
    DrawText = DrawTextW(hdc, StrPtr(strStr), Len(strStr), lpRect, wFormat)
End Function

