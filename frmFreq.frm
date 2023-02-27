VERSION 5.00
Begin VB.Form frmFreq 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ViStart_FrequentPrograms"
   ClientHeight    =   4575
   ClientLeft      =   195
   ClientTop       =   8610
   ClientWidth     =   3690
   ClipControls    =   0   'False
   FillColor       =   &H0080FFFF&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   246
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRollover 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   540
      Left            =   0
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   246
      TabIndex        =   0
      Top             =   -60
      Width           =   3690
   End
   Begin VB.Line theSeperator 
      BorderColor     =   &H00F5E4D6&
      Visible         =   0   'False
      X1              =   11
      X2              =   240
      Y1              =   76
      Y2              =   76
   End
End
Attribute VB_Name = "frmFreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_jumpListButton As pcMemDC
Private m_jumpListArrow As pcMemDC
Private m_rolloverBmp As pcMemDC

Private m_addToViPadCommand As String
Private m_viPadInstalled As Boolean

Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long

Private Type LIST_ITEM2
    Text As String
    SubText As String
    Shell As String
    Top As Long
    'Icon As DEFICON
End Type

Private Type LIST_ITEM
    Icon As ViIcon
    Text As String
    Shell As String
    Top As Long
    MRUList As JumpList
End Type

Private lstItems() As LIST_ITEM
Private iCurIndex As Long

Private m_selfRect As RECT

Private bKeyboardMode As Boolean

Private m_intIconSize As Integer
Private m_lngTopMargin As Long
Private m_trackingMouse As Boolean

Private Const m_cLeftMargin As Integer = 7

Private m_maxItems As Long

Public Event onNotifyAllPrograms()
Public Event onExitSide(ByVal index As Long)
Public Event onRequestCloseStartMenu()
Public Event onRequestShowJumpList(ByRef bSuccess As Boolean, ByRef theJumpList As JumpList)


Private m_theFont As ViFont
Private m_clientColour As Long
Private m_itemCap As Long
Private m_jumpPinnedItemIndex As Long
Private m_lastSelectedFile As String

Private m_HasJumpList As Boolean
Private m_shellIconSize As Long
Private m_dragCounter As Long

Private WithEvents m_vistaMenu As frmVistaMenu
Attribute m_vistaMenu.VB_VarHelpID = -1
Private WithEvents ProgramsDBEvents As clsProgramDB
Attribute ProgramsDBEvents.VB_VarHelpID = -1

Implements IHookSink

Public Function RolloverWithKeyboard(theRequestedIndex As Long)
    bKeyboardMode = True
    picRollover.Visible = True
    
    Rollover theRequestedIndex
End Function

Private Function ItemGap() As Long
    ItemGap = (m_intIconSize + 6)
End Function

Public Property Let ClientColour(newColour As Long)
    m_clientColour = newColour
End Property

Public Property Let TrueFont(newFont As ViFont)

    Set m_theFont = newFont
    
    picRollover.fontSize = newFont.Size
    Me.ForeColor = newFont.Colour

End Property

Sub SelectTopMost()
    picRollover.Visible = True

    iCurIndex = LBound(lstItems)
    RolloverWithKeyboard LBound(lstItems)
End Sub

Sub SelectBottomMost()
    picRollover.Visible = True

    If m_itemCap > 0 Then
        iCurIndex = m_itemCap
        RolloverWithKeyboard m_itemCap
    End If

End Sub

Sub RequestExecuteSelected()
    ExecuteSelected
End Sub

Sub ResetRollover()

    'Hide rollover object
    picRollover.Top = -picRollover.Height
    m_jumpPinnedItemIndex = -1
    iCurIndex = -1
    
    PaintItems
End Sub

Private Sub PopulateItemsFromCollection(ByVal slotIndexStart As Long, ByRef topValue As Long, ByRef sourceCollection As Collection)
    
    Debug.Print "PopulateItemsFromCollection:: " & slotIndexStart & " - " & topValue
    On Error GoTo Handler
    
Dim itemIndex As Long
Dim appDescription As String
Dim theCap As Long
Dim resolvedLnk As String
Dim thisProgram As clsProgram

    slotIndexStart = slotIndexStart - 1

    For itemIndex = 1 To sourceCollection.count
    
        If (slotIndexStart + itemIndex) < m_maxItems Then
            
            With lstItems(slotIndexStart + itemIndex)
                Set thisProgram = sourceCollection(itemIndex)
            
                'sourceCollection(itemIndex).Path
                appDescription = GetAppDescription(thisProgram.Path)
                
                If appDescription <> vbNullString And thisProgram.Caption = vbNullString Then
                    .Text = appDescription
                Else
                    .Text = VarScan(thisProgram.Caption)
                End If
                
                .Shell = thisProgram.Path
                Set .Icon = thisProgram.Icon
                'Set .Icon = IconManager.GetViIcon(.Shell, True)
                '.Icon.LoadIconFromFile .Shell, True
                Set .MRUList = GetImageJumpList(ResolveLink(.Shell))

                .Top = topValue
            End With
        End If
    
        topValue = topValue + ItemGap
    Next
Exit Sub
Handler:
    MsgBox Err.Description

    LogError "PopulateItemsFromCollection:: " & Err.Description, Me.Name
End Sub

Sub PopulateItems()
    On Error GoTo Handler

Dim itemIndex As Long
Dim actualProgramIndex As Long

Dim bExit As Boolean
Dim topValue As Long
    
    'Calculate Item Max
    If Settings.Programs.TotalProgramCount < m_maxItems Then
        m_itemCap = Settings.Programs.TotalProgramCount - 1
    Else
        m_itemCap = m_maxItems
    End If
    
    If m_itemCap = -1 Then
        Exit Sub
    End If
    
    topValue = m_lngTopMargin
    
    If Settings.Programs.PinnedPrograms.count > 0 Then
        theSeperator.Visible = True
    
        PopulateItemsFromCollection 0, topValue, Settings.Programs.PinnedPrograms
        MoveSeperator topValue - 1
        topValue = topValue + 5
    Else
        theSeperator.Visible = False
    End If
    
    PopulateItemsFromCollection Settings.Programs.PinnedPrograms.count, topValue, Settings.Programs.FrequentPrograms
    PaintItems
    
    Exit Sub
Handler:
    LogError "PopulateItems:: " & Err.Description, "frmFreq"
End Sub

Sub PaintItems()

    On Error GoTo Handler
    Me.Cls
    
Dim lngItemPaintIndex As Long
Dim RECT
Dim hasMRUList As Boolean
Dim dropY As Single

    For lngItemPaintIndex = 0 To (m_itemCap)
        Me.Font = OptionsHelper.PrimaryFont
        hasMRUList = Not lstItems(lngItemPaintIndex).MRUList.IsEmpty
    
        If m_jumpPinnedItemIndex = lngItemPaintIndex Then
            BitBlt Me.hdc, 1, lstItems(lngItemPaintIndex).Top - 2, m_rolloverBmp.Width, m_rolloverBmp.Height, m_rolloverBmp.hdc, 0, 0, vbSrcCopy
            BitBlt Me.hdc, 1 + picRollover.ScaleWidth - (m_jumpListButton.Width / 2), lstItems(lngItemPaintIndex).Top - 2, m_jumpListButton.Width / 2, m_jumpListButton.Height, m_jumpListButton.hdc, 0, 0, vbSrcCopy
        End If
        
        Me.ForeColor = m_theFont.Colour
        DrawTextMe Me, lstItems(lngItemPaintIndex).Text, m_cLeftMargin + m_intIconSize, lstItems(lngItemPaintIndex).Top, 9
        
        'Draw an icon if there is any
        If Not lstItems(lngItemPaintIndex).Icon Is Nothing Then lstItems(lngItemPaintIndex).Icon.DrawIconEx Me.hdc, 3, lstItems(lngItemPaintIndex).Top, 32, 32

        If lngItemPaintIndex >= Settings.Programs.PinnedPrograms.count And hasMRUList Then
            dropY = 0.5
        Else
            dropY = 0
        End If
        
        If hasMRUList And Not m_jumpPinnedItemIndex = lngItemPaintIndex Then
            BitBlt Me.hdc, picRollover.ScaleWidth - (m_jumpListButton.Width / 4) - (m_jumpListArrow.Width / 2) + 2, lstItems(lngItemPaintIndex).Top + ((m_jumpListButton.Height / 2) - (m_jumpListArrow.Height / 2)) - (2 + dropY), m_jumpListArrow.Width, m_jumpListArrow.Height, m_jumpListArrow.hdc, 0, 0, vbSrcCopy
        End If
    Next
    
    Exit Sub
Handler:
    LogError "PaintItems:: " & Err.Description, "frmFreq"
End Sub

Private Sub DrawTextMe(ObjSender As Object, sText As String, X As Long, Y As Long, lngSize As Long)

    On Error GoTo Handler

Dim lSuccess As Long
Dim lngHdc As Long
Dim myRect As RECT
Dim rSize As RECT

    lngHdc = ObjSender.hdc

    ObjSender.fontSize = m_theFont.Size

    'Do we need 1 line or 2
    rSize.Right = 200
    
    ObjSender.FontName = m_theFont.Face

    lSuccess = DrawTextW(lngHdc, StrPtr(sText), Len(sText), _
    rSize, DT_CALCRECT Or DT_WORDBREAK Or DT_NOPREFIX)
    
    myRect.Left = X
    
    If rSize.Bottom < (m_intIconSize / 2) Then
        'Single Line
        'myRect.Top = y + (m_intIconSize / 2) - 8
        myRect.Top = Y + 8
    Else
        'Double Line
        myRect.Top = Y + (m_intIconSize / 3.5) - 8
    End If
    
    myRect.Right = X + 200
    myRect.Bottom = Y + m_intIconSize
    
    lSuccess = DrawTextW(lngHdc, StrPtr(sText), Len(sText), _
    myRect, DT_WORD_ELLIPSIS Or DT_WORDBREAK Or DT_NOPREFIX)

    Exit Sub
Handler:
    LogError Err.Description, Me.Name & "::DrawTextMe"

End Sub

Private Sub Form_Initialize()
    m_lngTopMargin = 2

    Set m_theFont = New ViFont
    Set m_jumpListButton = New pcMemDC
    Set m_jumpListArrow = New pcMemDC
    Set m_rolloverBmp = New pcMemDC
    Set ProgramsDBEvents = Settings.Programs
    
    picRollover.Top = -picRollover.Height
    
    m_jumpPinnedItemIndex = -1
    m_shellIconSize = GetIconDimensions
End Sub

Private Sub Form_Load()
    Call HookWindow(Me.hWnd, Me)
    
    m_intIconSize = 32
    SetControlProperties
       
    picRollover.fontSize = 9
    
    Me.ForeColor = RGB(40, 40, 40)
    m_clientColour = RGB(70, 70, 70)
    
    PopulateItems
    Exit Sub
Handler:
    CreateError "frmFreq", "Form_Load()", Err.Description
End Sub

Private Sub SetControlProperties()
    
Dim returnBitmap As IPictureDisp
    
    m_jumpListButton.CreateFromPicture GetResourceBitmap("JUMPLIST_BUTTON_32")
    m_jumpListArrow.CreateFromPicture GetResourceBitmap("JUMPLIST_ARROW_32")
    m_rolloverBmp.CreateFromPicture GetResourceBitmap("FREQ_ROLLOVER_32")
    Set picRollover.Picture = GetResourceBitmap("FREQ_ROLLOVER_32")

    theSeperator.BorderColor = Layout.FrequentProgramsSeperatorColour

    Me.BackColor = g_sVar_Layout_BackColour
    Me.FontItalic = False
    
    picRollover.BackColor = g_sVar_Layout_BackColour
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'TestRolloverVisability
    
    'If IsAnyFileMenuShowing Then Exit Sub

Static lastPosition As points
Static lastButton As Integer
    
    If bKeyboardMode Then
        bKeyboardMode = False
        
        Debug.Print "HARE!"
        picRollover.Visible = False
    End If
    
    If lastPosition.X = X And lastPosition.Y = Y And lastButton = Button Then
        Exit Sub
    End If
    
    lastButton = Button
    
    lastPosition.X = X
    lastPosition.Y = Y
    
    If m_trackingMouse = False Then

        m_viPadInstalled = IsViPadInstalled
        m_trackingMouse = TrackMouse(Me.hWnd)
        picRollover.Visible = True
        Debug.Print "Visible!!"
    End If
    
    If m_vistaMenu Is Nothing Then UpdateRolloverStatus CreatePointL(CLng(Y), CLng(X))
End Sub

Private Sub Form_Resize()
    MaxItems = Floor(Me.ScaleHeight / ItemGap)
    
    theSeperator.X2 = Me.ScaleWidth - 7
End Sub

Private Property Let MaxItems(newMaxItems As Long)

    If newMaxItems <> m_maxItems Then
        m_maxItems = newMaxItems
        
        ReDim lstItems(m_maxItems)
        PopulateItems
    End If

End Property

Public Sub TestRolloverVisability()
    If Not m_vistaMenu Is Nothing Then Exit Sub

Dim cursorPos As win.POINTL
    
    GetCursorPos cursorPos
    ScreenToClient Me.hWnd, cursorPos
    
    If cursorPos.X > 0 And cursorPos.X < Me.ScaleWidth And _
        cursorPos.Y > 0 And cursorPos.Y < Me.ScaleHeight Then
        
        picRollover.Visible = True
    Else
        picRollover.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Settings.Programs.DumpPrograms
    If Not m_vistaMenu Is Nothing Then Unload m_vistaMenu
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
    On Error GoTo Handler
    
    If msg = WM_MOUSELEAVE Then
        m_trackingMouse = False
        
        Debug.Print "Leaving!"
        TestRolloverVisability
    Else
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = _
           CallOldWindowProcessor(hWnd, msg, wp, lp)
    End If
    
    Exit Function
Handler:
    Debug.Print Err.Description

    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       CallOldWindowProcessor(hWnd, msg, wp, lp)
End Function

Private Sub Form_KeyDown(lngKeyCode As Integer, Shift As Integer)

    If lngKeyCode = vbKeyRight Or lngKeyCode = vbKeyLeft Then
    
        picRollover.Top = -picRollover.Height
        
        'Set KeyBoard.mFocusObj = frmStartMenuBase.Keyboad_CallBack
        
        RaiseEvent onExitSide(iCurIndex + 1)
        Exit Sub
    End If

    If lngKeyCode = vbKeyDown Then
        'Down
    
        If iCurIndex = m_itemCap Then
            picRollover.Top = -picRollover.Height
            RaiseEvent onNotifyAllPrograms
            Exit Sub
        Else
            iCurIndex = iCurIndex + 1
        End If
        
        RolloverWithKeyboard iCurIndex

        
    ElseIf lngKeyCode = vbKeyUp Then
        'Up

        bKeyboardMode = True

        If iCurIndex = 0 Then

            picRollover.Top = -picRollover.Height
            RaiseEvent onNotifyAllPrograms
            Exit Sub
        Else
            If iCurIndex = -1 Then
                'frmError.ShowError "frmFreq", "onKeyUp", "Index out of sync"
                
                'frmError.cmdTerminate.Enabled = True
                'frmError.cmdIgnore.Enabled = False
                
                iCurIndex = 1
                Exit Sub
            End If
            
            iCurIndex = iCurIndex - 1
        End If
        
        RolloverWithKeyboard iCurIndex
    End If

End Sub

Private Sub Keyboad_CallBack_onKeyUp(ByVal lngKeyCode As Long)

    If lngKeyCode = 13 Then
        ExecuteSelected
    End If

End Sub

Private Sub Keyboad_CallBack_onLostFocus()
    
    'Hide rollover object
    picRollover.Top = -picRollover.Height

End Sub

Private Function ValidCurIndex() As Boolean
    If iCurIndex >= LBound(lstItems) And iCurIndex <= UBound(lstItems) Then
        ValidCurIndex = True
    End If
End Function

Private Sub ShowSelectJumpList()
    On Error GoTo Handler

Dim bResult As Boolean

    If iCurIndex >= LBound(lstItems) And iCurIndex <= UBound(lstItems) Then

        RaiseEvent onRequestShowJumpList(bResult, lstItems(iCurIndex).MRUList)
        If bResult Then
    
            If m_jumpPinnedItemIndex = iCurIndex Then
                m_jumpPinnedItemIndex = -1
            Else
                m_jumpPinnedItemIndex = iCurIndex
            End If
            
            PaintItems
        
        End If
    End If
    
    Exit Sub
Handler:
End Sub

Private Sub ExecuteSelected()
    Settings.Programs.UpdateByProgramPath m_lastSelectedFile
    
    If m_lastSelectedFile = "!default_menu" Then
        ShowNormalWindowsMenu
    Else
        SelectBestExecutionMethod m_lastSelectedFile
    End If
    
    PopulateItems
    
    iCurIndex = -1
End Sub

Private Sub UpdateRolloverStatus(cPos As win.POINTL)

Dim iTop As Integer
Dim rectSelf As RECT

Dim i As Long
Dim bOnTopOfButton As Boolean
Dim suggestedIndex As Long

    'Debug.Print "UpdateRolloverStatus:: " & cPos.X & " - " & cPos.Y

    suggestedIndex = Floor(cPos.Y / ItemGap)
    
    If Settings.Programs.PinnedPrograms.count > 0 Then
        If suggestedIndex = Settings.Programs.PinnedPrograms.count Then
            Debug.Print "changing index!"
            suggestedIndex = Floor((cPos.Y - 10) / ItemGap)
        End If
    End If
    
    If (suggestedIndex > -1) And suggestedIndex <= m_itemCap Then
                
                
        'Cancel Keyboard Mode
        bKeyboardMode = False
        bOnTopOfButton = True

        If iCurIndex <> suggestedIndex Then
        
            
            Rollover suggestedIndex

            iCurIndex = suggestedIndex
            Debug.Print "01# Changing iCurIndex to " & suggestedIndex
            
            Debug.Print "So::" & lstItems(suggestedIndex).Text & " :: " & suggestedIndex & " - " & m_itemCap
        End If
    End If

    If Not bOnTopOfButton Then

        If Not bKeyboardMode Then

            iCurIndex = -1
            Debug.Print "02# Changing iCurIndex to " & iCurIndex
            
            'Hide rollover object
            picRollover.Top = -picRollover.Height
        End If
    End If

End Sub

Private Function Rollover(ByVal lngNewRolloverIndex As Long)
    On Error GoTo Handler
    
Dim hasMRUList As Boolean
    
    If lngNewRolloverIndex = -2 Then
        'Debug.Print "Bail 04"
        Exit Function
    End If
    
    If UBound(lstItems) = 0 Then
        'Debug.Print "Bail 00"
        Exit Function
    End If
    
    If Not m_vistaMenu Is Nothing Then
        'Debug.Print "Bail 01"
        Exit Function
    End If
    
    Debug.Print "Rollovering over:: " & lngNewRolloverIndex
    
    If lngNewRolloverIndex >= LBound(lstItems) And lngNewRolloverIndex <= UBound(lstItems) Then
    
        hasMRUList = Not lstItems(lngNewRolloverIndex).MRUList.IsEmpty
    
        Rollover = True
        
        picRollover.FontName = m_theFont.Face
        picRollover.Cls
        picRollover.ForeColor = m_theFont.Colour
        
        DrawTextMe picRollover, lstItems(lngNewRolloverIndex).Text, m_cLeftMargin + m_intIconSize - 1, 2, 9
        picRollover.Move 1, lstItems(lngNewRolloverIndex).Top - 2
        
        lstItems(lngNewRolloverIndex).Icon.DrawIconEx picRollover.hdc, 2, 2, 32, 32
        'ExtractIcon lstItems(lngNewRolloverIndex).Shell, picIconRollover, m_intIconSize
        
        'BitBlt picRollover.hdc, 2, 2, m_intIconSize, m_intIconSize, picIconRollover.hdc, 0, 0, vbSrcCopy
        'picIconRollover.Cls

        
        If hasMRUList Then
            'BitBlt picRollover.hdc, picRollover.ScaleWidth - m_jumpListButton.Width, 0, m_jumpListButton.Width, m_jumpListButton.Height, m_jumpListButton.hdc, 0, 0, vbSrcCopy
            
            BitBlt picRollover.hdc, picRollover.ScaleWidth - (m_jumpListButton.Width / 2), 0, m_jumpListButton.Width, m_jumpListButton.Height, m_jumpListButton.hdc, (m_jumpListButton.Width / 2), 0, vbSrcCopy
        End If
    End If
    
    Exit Function
Handler:
    CreateError "frmFreq", "Rollover", Err.Description
End Function

Private Function DrawJumpListButton(theState As Long)

    BitBlt picRollover.hdc, (picRollover.ScaleWidth - m_jumpListButton.Width) / 2, 0, m_jumpListButton.Width, m_jumpListButton.Height, m_jumpListButton.hdc, m_jumpListButton.Width / 2, 0, vbSrcCopy
    
End Function

Private Function JustExe(strPath As String) As String

Dim lngRightPos As Long
Dim sReturn As String

    On Error GoTo Handler

    strPath = VarScan(strPath)

    If Left$(strPath, 1) = """" Then
        lngRightPos = InStr(2, strPath, """")
        If lngRightPos > -1 Then
            sReturn = Mid$(strPath, 2, lngRightPos - 2)
        Else
            sReturn = ""
        End If
    Else
        sReturn = strPath
    End If
    
    JustExe = sReturn
    Exit Function
    
Handler:
    JustExe = ""

End Function

Private Sub m_vistaMenu_onClick(theItemTag As String)
    
Dim theCommand As String
Dim theParameter As String
Dim sP() As String

    sP = Split(theItemTag, "@")
    theCommand = sP(0)
    theParameter = sP(1)
    
    'MsgBox theCommand & ":" & theParameter
    m_vistaMenu_onInActive

    Select Case UCase$(theCommand)
    
    Case "COPYTOVIPAD"
        ShellEx m_addToViPadCommand
    
    Case "COPYTODESKTOP"
        If FSO.FolderExists(sVar_Reg_Desktop) And Not FSO.FileExists(sVar_Reg_Desktop & "\" & GetFileName(theParameter)) Then
            FSO.CopyFile theParameter, sVar_Reg_Desktop & "\" & GetFileName(theParameter), False
        End If
    
    Case "OPEN"
        RaiseEvent onRequestCloseStartMenu
        SelectBestExecutionMethod theParameter
    
    Case "RUNASADMIN"
        RaiseEvent onRequestCloseStartMenu
        
        ShellEx theParameter, "runas"
    
        Settings.Programs.UpdateByProgramPath theParameter
        PopulateItems
    
    Case "TOGGLEPIN"
        Settings.Programs.TogglePin theParameter
        
    Case "REMOVEITEM"
        Settings.Programs.RemoveItem theParameter
        
    Case "PROPERTIES"
        RaiseEvent onRequestCloseStartMenu
        ShellEx theParameter, "properties"
        
    Case Else
        'MsgBox theCommand
        
    End Select
        
    PopulateItems
    ResetRollover
End Sub

Private Sub m_vistaMenu_onInActive()
    Debug.Print "INACTIVE!"
    
    If Not m_vistaMenu Is Nothing Then
        Unload m_vistaMenu
        Set m_vistaMenu = Nothing
    End If
    
    
    TestRolloverVisability
End Sub

Private Function GetSelectedPath() As String
    If iCurIndex >= LBound(lstItems) And iCurIndex <= UBound(lstItems) Then
        GetSelectedPath = lstItems(iCurIndex).Shell
    End If
End Function

Private Sub picRollover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If X < (picRollover.ScaleWidth - (m_jumpListButton.Width / 2)) Then
            m_dragCounter = m_dragCounter + 1
            
            If m_dragCounter = 4 Then
                m_dragCounter = 0
                picRollover.OLEDrag
            End If
        End If
    End If
End Sub

Private Sub picRollover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Handler

    If Button = vbLeftButton Then

        If ValidCurIndex Then
            If X > (picRollover.ScaleWidth - (m_jumpListButton.Width / 2)) And _
                Not lstItems(iCurIndex).MRUList.IsEmpty Then 'Error not set
    
                ShowSelectJumpList
            Else
                m_lastSelectedFile = GetSelectedPath()
            
                RaiseEvent onRequestCloseStartMenu
                ExecuteSelected
            End If
        End If
        
    ElseIf Button = vbRightButton Then

        If iCurIndex < LBound(lstItems) Or iCurIndex > UBound(lstItems) Then
            Exit Sub
        End If

        If Not m_vistaMenu Is Nothing Then Unload m_vistaMenu
        Set m_vistaMenu = New frmVistaMenu
        
        m_vistaMenu.AddItem GetPublicString("strOpen"), "OPEN@" & lstItems(iCurIndex).Shell, True

        If LCase$(Right$(ResolveLink(lstItems(iCurIndex).Shell), 3)) = "exe" Then
            m_vistaMenu.AddItem GetPublicString("strRunAsAdmin"), "RUNASADMIN@" & lstItems(iCurIndex).Shell
        End If
        
        m_vistaMenu.AddItem ""
        
        If iCurIndex < Settings.Programs.PinnedPrograms.count Then
            m_vistaMenu.AddItem GetPublicString("strUnpinToStartMenu"), "TOGGLEPIN@" & lstItems(iCurIndex).Shell
        Else
            m_vistaMenu.AddItem GetPublicString("strPinToStartMenu"), "TOGGLEPIN@" & lstItems(iCurIndex).Shell
        End If
        m_vistaMenu.AddItem ""
        m_vistaMenu.AddItem GetPublicString("strRemoveFromList"), "REMOVEITEM@" & lstItems(iCurIndex).Shell
        
        If FileExists(lstItems(iCurIndex).Shell) Then
        
            m_vistaMenu.AddItem ""
        
            Dim lnkFileShellRegKey As RegistryKey: Set lnkFileShellRegKey = Registry.ClassesRoot.OpenSubKey("lnkfile\shell")

            If Not lnkFileShellRegKey.OpenSubKey("Add to ViPad") Is Nothing Then
                m_addToViPadCommand = lnkFileShellRegKey.OpenSubKey("Add to ViPad").GetValue("command", vbNullString)
                m_addToViPadCommand = Replace(m_addToViPadCommand, "%1", lstItems(iCurIndex).Shell)
                
                m_vistaMenu.AddItem GetPublicString("strCopyToViPad"), "COPYTOVIPAD@NULL"
            Else
                If m_viPadInstalled Then
                    m_addToViPadCommand = GenerateViPadAddToCommand(lstItems(iCurIndex).Shell)
                
                m_vistaMenu.AddItem GetPublicString("strCopyToViPad"), "COPYTOVIPAD@NULL"
            '    'Else
            '    '    m_addToViPadCommand = "http://lee-soft.com/vipad"
                End If

            End If
        
            m_vistaMenu.AddItem GetPublicString("strCopyToDesktop"), "COPYTODESKTOP@" & lstItems(iCurIndex).Shell
    
            m_vistaMenu.AddItem ""
            m_vistaMenu.AddItem GetPublicString("strProperties"), "PROPERTIES@" & lstItems(iCurIndex).Shell
        End If
        
        Debug.Print "Attemping Resurrection!"
        m_vistaMenu.Resurrect True
    End If

    Exit Sub
Handler:
    LogError Err.Description, "frmFreq"
End Sub

Sub MoveSeperator(theY As Long)

    theSeperator.Y1 = theY
    theSeperator.Y2 = theY
End Sub

Private Sub picRollover_OLESetData(Data As DataObject, DataFormat As Integer)
    On Error GoTo Handler
    Data.Files.Add lstItems(iCurIndex).Shell
Handler:
End Sub

Private Sub ProgramsDBEvents_onFrequentProgramsFlushed()
    PopulateItems
End Sub

Private Sub ProgramsDBEvents_onMetroShortcutAdded()
    PopulateItems
End Sub

Private Sub picRollover_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    On Error GoTo Handler
    ' Set data format to file.
    Data.SetData , vbCFFiles
    ' Display the move mouse pointer..
    AllowedEffects = vbDropEffectCopy
    
    Data.Files.Add lstItems(iCurIndex).Shell
Handler:
End Sub

Private Sub ProgramsDBEvents_onRequestRedraw()
    PopulateItems
End Sub
