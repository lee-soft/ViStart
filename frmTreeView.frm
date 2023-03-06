VERSION 5.00
Begin VB.Form frmTreeView 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Vista_ProgramMenu"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   612
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFocusGrab 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   615
      Left            =   -750
      TabIndex        =   0
      Top             =   2460
      Width           =   735
   End
   Begin VB.VScrollBar scrVerticle 
      Height          =   870
      Left            =   0
      Max             =   0
      SmallChange     =   15
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DISPLAY_MAX As Long = 1024
Private Const SCROLL_CHANGE As Long = 38

Private WithEvents mTreeViewData As clsTreeview
Attribute mTreeViewData.VB_VarHelpID = -1
Private WithEvents mTreeViewSearchResults As clsTreeview
Attribute mTreeViewSearchResults.VB_VarHelpID = -1
Private WithEvents mSearchProvider As frmSearchMaster
Attribute mSearchProvider.VB_VarHelpID = -1


Private Const M_CBORDERSIZE_NOSCROLL As Long = 5
Private Const M_CROLLOVER_LEFTSTART_NOSCROLL As Long = 5
Private Const MC_SEPERATOR_NODE As Integer = -1

Private Const M_SEPARATOR_GAP As Long = 12
Private Const m_cNodeSpace As Long = 19

'Private m_dcPrograms As pcMemDC
'Private m_dcFiles As pcMemDC
Private m_dcLib As Collection

Private m_dcTreeSeparator As pcMemDC
Private m_dcIndexIcons As pcMemDC
Private m_dcSearchIcon As pcMemDC

Private m_nShowAllResults As INode

Private m_colTypes As Collection

Private Const m_cIconSize As Long = 16
Private Const m_cBorderSize As Long = 4

Private m_textSize As Long

Private m_seperator() As Long
Private m_seperator_index As Long
Private m_seperator_inuse() As Boolean
Private m_Seperator_fontColour As Long

Private m_bRestrictSearch As Boolean
Private m_bExceedsLimits As Boolean

Private m_recBlank As RECT
Private m_paBlank As POINTL

Private m_cRecNoResults As RECT
Private m_ProgramsBitmapHeight As Long

Private mDx As ISoftX
Private mrTextPos As RECT

Private mvRolloverPos As POINTL
Private mvRolloverDesc As POINTL

Private m_paSearchIconDestinationPosition As POINTL
Private m_paSearchIconSourcePosition As POINTL

Private mvIconPos As POINTL
Private mvIconIndex As POINTL

Private mStdFont As StdFont

Private mLngLevel As Long

Private mlngLowerBound As Long

Private mLngTopStart As Long
Private mLngLeftStart As Long

Private mLngRounded As Long

Private mcVisibleNodes As Collection

Private mstrKeyWord As String
Private mstrOldKeyWord As String
Private bKeyWordChanged As Boolean

Private m_nTargetNode As INode
Private m_nLastSelectedNode As INode

Private m_dcRollover As pcMemDC

Private m_rTextRolloverPos As RECT
Private mvRolloverIconPos As POINTL
Private mvRolloverIconIndex As POINTL

Private mSearchResultsNode As INode
Private m_bSearchMode As Boolean
Private m_bKeyboardMode As Boolean

Private m_rrolloverSize As RECT
Private m_lngNodeIndex As Long

Private m_rShowAllResults As RECT

Private m_bSearchResultsEmpty As Boolean
Private m_bForceRepaint As Boolean

Private m_TextDisplayWidth As Long

Private m_pCursor As POINTL
Private m_backBrush As GDIBrush

Private WithEvents m_contextMenu As frmVistaMenu
Attribute m_contextMenu.VB_VarHelpID = -1

Private lngLastPosition As Long

Private m_trackingMouse As Boolean
Private m_nodeDisplayLimit As Long

Private m_separatorFont As StdFont
Private m_originalTextColour As Long
Private m_lastTypeCount As Long
Private m_programCount As Long
Private m_queryType As TV_Type
Private m_addToViPadCommand As String
Private m_viPadInstalled As Boolean

Public Event onClick(srcNode As INode)
Public Event onNotifyAllPrograms()
Public Event onExit(ByVal index As Long)
Public Event onRequestRecentProgramsRefresh()
Public Event onRequestCloseStartMenu()

Implements IHookSink

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    Set Logger = m_logger
End Property

Public Function callback_ForceRepaint(newNode As INode)
    m_queryType.Children.Add newNode

    mSearchProvider_onNewItem
End Function

Public Property Let TrueFont(newFont As GDIFont)
    With mStdFont
        .Name = newFont.FontFace
        .Size = newFont.FontWeight
    End With
    
    m_separatorFont.Name = newFont.FontFace
    m_separatorFont.Size = 20

    mDx.Font = mStdFont
End Property

Public Property Let SeparatorFontColour(newColour As Long)
    m_Seperator_fontColour = newColour
End Property

Public Property Let BackColour(newColour As Long)
    m_backBrush.Colour = newColour
End Property

Public Property Let ForeColour(newColour As Long)
    mDx.TextColour = newColour
    m_originalTextColour = newColour
End Property

Public Sub RequestActionNode()

    If Not m_bSearchResultsEmpty Then
        ActionTargetNode
    End If

End Sub

Public Function ColapseAllAndResetStrictSearch()
    mTreeViewData.ColapseAll
    m_bRestrictSearch = True
End Function

Public Function ResetScrollbarsValues()
    scrVerticle.Value = 0
End Function

Public Function ResetKeyboardStatus()
    m_bKeyboardMode = False
    Form_Paint
End Function

Public Property Get RootNode() As INode
    Set RootNode = mTreeViewData.RootNode
End Property

Public Property Get RootPrograms() As Collection
    Set RootPrograms = mTreeViewData.AllNodes
End Property

Public Function AddType(ByRef srcTV_Type As TV_Type, Optional desiredPosition As Long = 0)
    
Dim newINode As New INode
Dim thisType As TV_Type

    If ExistInCol(m_colTypes, srcTV_Type.Caption) Then
        Exit Function
    End If

    With newINode
        .Caption = srcTV_Type.Caption
        
        'Set as Seperator
        .IconPosition = MC_SEPERATOR_NODE

        .GetSpaces .Caption
        .Width = CalcNodeWidth(.Caption)
    End With
    
    LoadBitmapType srcTV_Type.Caption
    
    Set srcTV_Type.Node = newINode

    If desiredPosition = 0 Or desiredPosition > m_colTypes.count Then
        m_colTypes.Add srcTV_Type, srcTV_Type.Caption
    Else
        m_colTypes.Add srcTV_Type, srcTV_Type.Caption, desiredPosition
    End If
    
    ReDim Preserve m_seperator(m_colTypes.count - 1)
    ReDim Preserve m_seperator_inuse(m_colTypes.count - 1)

    ReCalculateVisibleNodes
    For Each thisType In m_colTypes
        thisType.DisplayLimit = m_nodeDisplayLimit
    Next
End Function

Public Function AddNode(strCaption As String, lngIconIndex As Long, strTag As String, Optional nParentNode As INode) As INode

Dim new_INode As INode
Dim rNodeRect As RECT
    
    Set new_INode = mTreeViewData.createNode(nParentNode)
    rNodeRect = mDx.GetTextRect(strCaption)
    
    With new_INode
        .Caption = strCaption
        .IconPosition = lngIconIndex
        .Width = rNodeRect.Right
        .Tag = strTag
    End With
    
    Set AddNode = new_INode
    
End Function

Public Function CalcNodeWidth(strCaption As String) As Long
    CalcNodeWidth = mDx.GetTextRect(strCaption).Right
End Function

Public Property Let Filter(new_strKeyWord As String)
    
    mstrKeyWord = new_strKeyWord
    mLngTopStart = -scrVerticle.Value + m_cBorderSize
    
    If new_strKeyWord = mstrOldKeyWord Then
        bKeyWordChanged = False
    Else
        mstrOldKeyWord = new_strKeyWord
        bKeyWordChanged = True
        
        
        
        mSearchProvider.ExecuteNewQuery new_strKeyWord
        mSearchProvider_onNewItem
        
        'If m_bKeyboardMode Then
                'If m_bSearchResultsEmpty Then
            SelectFirstItem
                'End If
        'Else
            'm_pCursor.X = -1
            'm_lngNodeIndex = -1
            
            'mLngRounded = m_cNodeSpace * m_lngNodeIndex

        'End If
        
    End If
    
    Form_Paint
    UpdateRolloverPosition
    
End Property

Function ShowContextMenu() As Boolean
    On Error GoTo Handler

    If (m_nTargetNode Is Nothing) Then
        ShowContextMenu = False
        Exit Function
    End If
    
    Set m_nLastSelectedNode = m_nTargetNode

    If Not m_contextMenu Is Nothing Then Unload m_contextMenu
    Set m_contextMenu = New frmVistaMenu
    
    If Not m_nTargetNode.IsFile Then
        
                m_contextMenu.AddItem GetPublicString("strExplore"), "EXPLORE", True
                
        If m_nTargetNode.Expanded Then
            m_contextMenu.AddItem GetPublicString("strCollapse"), "COLLAPSE", True
        Else
            m_contextMenu.AddItem GetPublicString("strExpand"), "EXPAND"
        End If
    Else
        m_contextMenu.AddItem GetPublicString("strOpen"), "OPEN", True
        
        m_contextMenu.AddItem GetPublicString("strRunAsAdmin"), "RUNASADMIN"
        m_contextMenu.AddItem ""
    End If
    
    If Settings.Programs.ExistsInPinned(m_nTargetNode.Tag) Then
        m_contextMenu.AddItem GetPublicString("strUnpinToStartMenu"), "TOGGLEPIN"
    Else
        m_contextMenu.AddItem GetPublicString("strPinToStartMenu"), "TOGGLEPIN"
    End If

    If FileExists(m_nTargetNode.Tag) Then
        m_contextMenu.AddItem ""
        
        Dim lnkFileRegKey As RegistryKey: Set lnkFileRegKey = Registry.ClassesRoot.OpenSubKey("lnkfile\shell")
        
        If Not lnkFileRegKey.OpenSubKey("Add to ViPad") Is Nothing Then
            m_addToViPadCommand = lnkFileRegKey.OpenSubKey("Add to ViPad").GetValue("command", vbNullString)
            m_addToViPadCommand = Replace(m_addToViPadCommand, "%1", m_nTargetNode.Tag)
            
            m_contextMenu.AddItem GetPublicString("strCopyToViPad"), "COPYTOVIPAD"
        Else
            If m_viPadInstalled Then
                m_addToViPadCommand = GenerateViPadAddToCommand(m_nTargetNode.Tag)
           
                m_contextMenu.AddItem GetPublicString("strCopyToViPad"), "COPYTOVIPAD"
        '    'Else
        '    '    m_addToViPadCommand = "http://lee-soft.com/vipad"
            End If
                
        End If
        
        m_contextMenu.AddItem GetPublicString("strCopyToDesktop"), "COPYTODESKTOP"
    
        m_contextMenu.AddItem ""
        m_contextMenu.AddItem GetPublicString("strProperties"), "PROPERTIES"
    End If
    
    m_contextMenu.Resurrect True, Me
    Exit Function
    
Handler:
    Logger.Error Err.Description, "ShowContextMenu"
End Function

Private Sub Form_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)
    
Dim hasPrograms As Boolean

    Set m_dcLib = New Collection
    Set m_backBrush = New GDIBrush

    Set m_nShowAllResults = New INode
    
    With m_nShowAllResults
        .Caption = UserVariable("strSeeAllResults")
        .Tag = "//SYS_SHOW_ALL"
    End With

    Set mStdFont = New StdFont
    Set m_separatorFont = New StdFont
        Set m_dcSearchIcon = New pcMemDC
    
    Set mDx = New ISoftX
    
    Set mTreeViewData = New clsTreeview
    Set mTreeViewSearchResults = New clsTreeview
    
    Set m_colTypes = New Collection
    
    Set m_dcRollover = New pcMemDC
    Set m_dcIndexIcons = New pcMemDC
    
    Set mSearchResultsNode = New INode
    
    m_Seperator_fontColour = RGB(30, 50, 135)
    m_backBrush.Colour = vbWhite
    
    mDx.TextColour = vbBlack
    m_originalTextColour = vbBlack
    
    m_dcRollover.CreateFromPicture GetResourceBitmap("ROLLOVER")
    'm_dcTreeSeparator.CreateFromPicture GetResourceBitmap("TREE_SEPARATOR")
    
    Set m_dcTreeSeparator = New pcMemDC
    m_dcTreeSeparator.CreateFromPicture GetResourceBitmap("TREE_SEPARATOR")

    'Set thisBitmap = GetResourceBitmap("programs")
    'If Not thisBitmap Is Nothing Then
        'm_displayTextInsteadOfBitmap = False
        'm_dcPrograms.CreateFromPicture thisBitmap
    'Else
        'm_dcPrograms.CreateFromPicture GetResourceBitmap("TREE_SEPARATOR")
    'End If
    
    'Set thisBitmap = GetResourceBitmap("files")
    'If Not thisBitmap Is Nothing Then
        'm_displayTextInsteadOfBitmap = False
        'm_dcFiles.CreateFromPicture thisBitmap
    'Else
        'm_dcFiles.CreateFromPicture GetResourceBitmap("TREE_SEPARATOR")
    'End If
    
    m_dcIndexIcons.CreateFromPicture GetResourceBitmap("INDEXED_ICON")
    m_dcSearchIcon.CreateFromPicture GetResourceBitmap("SEARCH_ICON")
    
    m_cRecNoResults.Top = m_dcTreeSeparator.Height
    m_cRecNoResults.Bottom = m_cRecNoResults.Top + 50
    
    m_cRecNoResults.Left = 20
    m_cRecNoResults.Right = m_TextDisplayWidth
    
    'Normal Node Space
    mLngTopStart = m_cBorderSize
    
    SetupTypes
End Sub

Private Function LoadBitmapType(szFileName As String)

Dim thisDc As New pcMemDC
Dim thisBitmap As IPictureDisp

    Set thisBitmap = GetResourceBitmap(szFileName)
    If Not thisBitmap Is Nothing Then
        thisDc.CreateFromPicture thisBitmap
        If Not ExistInCol(m_dcLib, szFileName) Then m_dcLib.Add thisDc, szFileName
    End If

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'RaiseEvent onKeyDown(KeyCode)

    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        'Form_Paint

        'frmStartMenuBase.Keyboad_CallBack.KeyListIndex = 0

        'KeyBoard.mFocusObj.onLostFocus
        'Set KeyBoard.mFocusObj = frmStartMenuBase.Keyboad_CallBack
        'KeyBoard.mFocusObj.onGotFocus

        'Exit Sub
        
        m_bKeyboardMode = False
        RaiseEvent onExit(m_lngNodeIndex - 1)

        Form_Paint
        Exit Sub
    End If

    If KeyCode = vbKeyDown Then
        moveRolloverDown
    ElseIf KeyCode = vbKeyUp Then
        moveRolloverUp
    ElseIf KeyCode = vbKeyReturn Then
        ActionTargetNode
    End If

    UpdateRolloverPosition

End Sub

Private Sub Form_Load()
    HookWindow Me.hWnd, Me
    
    m_pCursor.X = -1
    
    With mStdFont
        .Name = sVar_sFontName
        .Size = 9
    End With
    
    With m_separatorFont
        .Name = sVar_sFontName
        .Size = 12
    End With

    mDx.hWnd = hWnd
    mDx.Font = mStdFont
    
    'picList_SubClass.SubClass Me.hWnd
    mTreeViewSearchResults.RootNode.Caption = GetPublicString("strPrograms")
    
    m_bRestrictSearch = True
End Sub

Sub ReCalculateVisibleNodes()

Dim totalNodesVisible As Long
Dim thisType As TV_Type

    totalNodesVisible = ((Me.ScaleHeight - (M_SEPARATOR_GAP * m_colTypes.count)) / 23.6666666666667)
    m_nodeDisplayLimit = Floor(totalNodesVisible / m_colTypes.count)
    
End Sub

Sub SetupTypes()

Dim m_TVTType As TV_Type

    Set m_TVTType = New TV_Type

    m_TVTType.Caption = GetPublicString("strPrograms")
    m_TVTType.DisplayLimit = 10
    m_TVTType.AllowQuery = True
    
    Set m_TVTType.Children = mTreeViewData.AllNodes

    AddType m_TVTType
    
    '
    If Not g_WDSInitialized Then
    
        Set m_TVTType = New TV_Type
        
        m_TVTType.Caption = GetPublicString("strFiles")
        m_TVTType.DisplayLimit = 5
        m_TVTType.AllowQuery = True
        
        Set m_TVTType.Children = g_colSearch
        
        
        AddType m_TVTType
    End If
    
    Set mSearchProvider = New frmSearchMaster
    Load mSearchProvider
End Sub

Private Function CalculateRolloverYFromCurrentIndex()

Dim seperatorIndex As Long
Dim selectedSeperator As Long
Dim nodeIndex As Long
Dim addition As Long
Dim seperatorIncrease As Long

    If m_lngNodeIndex > mcVisibleNodes.count Then
        CalculateRolloverYFromCurrentIndex = 0
        Exit Function
    End If
    
    If Len(mstrKeyWord) > 0 Then
    
        If m_lngNodeIndex > 1 And m_lngNodeIndex < mcVisibleNodes.count Then
            If mcVisibleNodes(m_lngNodeIndex).IconPosition = MC_SEPERATOR_NODE Then
                CalculateRolloverYFromCurrentIndex = -1
                Exit Function
            End If
        End If
    
        For nodeIndex = 1 To m_lngNodeIndex
            If mcVisibleNodes(nodeIndex).IconPosition = MC_SEPERATOR_NODE Then
                'If m_colTypes(mcVisibleNodes(nodeIndex).Caption).Children.count > 0 Then
                    selectedSeperator = selectedSeperator + 1
                'End If
            End If
        Next
        selectedSeperator = selectedSeperator - 1

        seperatorIncrease = (selectedSeperator + 1) * M_SEPARATOR_GAP 'Difference between node height and seperater total height is 2
    End If
    
    addition = 17
    CalculateRolloverYFromCurrentIndex = ((m_lngNodeIndex * m_cNodeSpace) + (seperatorIncrease) - addition) - Me.scrVerticle.Value
End Function

Private Function ValuePlusScrollbar(sourceValue As Long) As Long
    ValuePlusScrollbar = sourceValue + scrVerticle.Value
End Function

Private Function CalculateInvisibleNodes()

Dim nodeIndex As Long

    For nodeIndex = 1 To m_lngNodeIndex
        If mcVisibleNodes(nodeIndex).IconPosition = MC_SEPERATOR_NODE Then
            If m_colTypes(mcVisibleNodes(nodeIndex).Caption).Children.count Then
                CalculateInvisibleNodes = CalculateInvisibleNodes + 1
            End If
        End If
    Next

End Function

Private Function FindNodeIndex(ByVal sourceY As Long)

Dim seperatorIndex As Long
Dim selectedSeperator As Long
Dim seperatorIncrease As Long
Dim invisibleSeperator As Long

    sourceY = sourceY - 2
    sourceY = ValuePlusScrollbar(sourceY)

    If Len(mstrKeyWord) > 0 Then
        For seperatorIndex = LBound(m_seperator) To UBound(m_seperator)
            If m_seperator(UBound(m_seperator) - seperatorIndex) <> -1 Then
                If sourceY > ValuePlusScrollbar(m_seperator(UBound(m_seperator) - seperatorIndex)) Then
                    selectedSeperator = (UBound(m_seperator) - seperatorIndex)
                    Exit For
                End If
            End If
        Next
        
        seperatorIncrease = (selectedSeperator + 1) * M_SEPARATOR_GAP 'Difference between node height and seperater total height is 2

        If sourceY > ValuePlusScrollbar(m_seperator(selectedSeperator)) And sourceY < ValuePlusScrollbar(m_seperator(selectedSeperator) + (m_cNodeSpace + M_SEPARATOR_GAP)) Then
            FindNodeIndex = -1
            Exit Function
        End If
    End If

    sourceY = sourceY - seperatorIncrease 'Difference between node height and seperater total height is 2
    FindNodeIndex = Floor(sourceY / m_cNodeSpace) + 1
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim newNodeIndex As Long
Dim originalY As Long

    If Not m_contextMenu Is Nothing Then
        Exit Sub
    End If

    If m_trackingMouse = False Then
        m_viPadInstalled = IsViPadInstalled
        m_trackingMouse = TrackMouse(Me.hWnd)
    End If

    If (m_pCursor.X = -1) Then
        m_pCursor.X = X
        m_pCursor.Y = Y
        
        Exit Sub
    Else
        If (m_pCursor.X <> X Or m_pCursor.Y <> Y) Then
            m_pCursor.X = X
            m_pCursor.Y = Y
        Else
            Exit Sub
        End If
    End If

    If Not m_bKeyboardMode Then
        m_bKeyboardMode = True
    End If

    newNodeIndex = FindNodeIndex(Y)
    originalY = Y
    
    Y = scrVerticle.Value + Y - m_cBorderSize + 1
    mLngRounded = RoundIt(Y + 9, m_cNodeSpace)
    
    If newNodeIndex <> m_lngNodeIndex Then
    
        If newNodeIndex > mcVisibleNodes.count Then
            newNodeIndex = mcVisibleNodes.count
        End If
    
        If newNodeIndex > 0 Then
    
            If Not mcVisibleNodes(newNodeIndex).IconPosition = MC_SEPERATOR_NODE Then
                m_lngNodeIndex = newNodeIndex
                mvRolloverPos.Y = CalculateRolloverYFromCurrentIndex

                UpdateRolloverPosition
            End If
        End If
    End If

Handler:
End Sub

Sub UpdateRolloverPosition()

    On Error GoTo Trapper
    
Dim new_nTargetNode As INode

    If (m_lngNodeIndex <= mcVisibleNodes.count) And _
        (m_lngNodeIndex > 0) Then
        
        Set new_nTargetNode = mcVisibleNodes(m_lngNodeIndex)
        'Debug.Print "#1 Setting new target node: " & m_lngNodeIndex
        
        If m_nTargetNode Is Nothing Then
            Set m_nTargetNode = mcVisibleNodes(m_lngNodeIndex)
            'Debug.Print "#2 Setting new target node: " & m_lngNodeIndex
        ElseIf Not (new_nTargetNode Is m_nTargetNode) Then
            Set m_nTargetNode = new_nTargetNode
            'Debug.Print "#3 Setting new target node: " & m_lngNodeIndex
        End If
        
        MoveRollover
    End If
    
    Exit Sub
Trapper:
    Logger.Error Err.Description, "UpdateRolloverPosition"

End Sub

Private Sub ResetSeperatorPositions()
    
Dim seperatorIndex As Long
    

    For seperatorIndex = LBound(m_seperator) To UBound(m_seperator)
        m_seperator(seperatorIndex) = -1
        m_seperator_inuse(seperatorIndex) = False
        
    Next
    
End Sub

Private Sub MoveRollover()

    On Error GoTo Reazon
    
Dim seperatorIndex As Long
    
    If m_nTargetNode Is Nothing Then
        Exit Sub
    End If
    
    If Len(mstrKeyWord) > 0 Then
        If m_nTargetNode.Tag = "//SYS_SHOW_ALL" And m_bKeyboardMode Then
            m_paSearchIconSourcePosition.X = 16
        
            mvRolloverPos.Y = m_rShowAllResults.Top - 2
            mvRolloverPos.X = m_rShowAllResults.Left
        Else
            m_paSearchIconSourcePosition.X = 0
        End If
    End If

    m_rTextRolloverPos.Top = mvRolloverPos.Y + 2
    m_rTextRolloverPos.Bottom = mvRolloverPos.Y + 20
    
    m_rTextRolloverPos.Left = m_nTargetNode.Left
    m_rTextRolloverPos.Right = m_TextDisplayWidth
    
    mvRolloverIconPos.Y = m_rTextRolloverPos.Top
    mvRolloverIconPos.X = m_nTargetNode.Left - m_cNodeSpace
    
    mvRolloverIconIndex.Y = m_nTargetNode.IconPosition * m_cIconSize

    Form_Paint
    
    Exit Sub
Reazon:
    Logger.Error Err.Description, "MoveRollover"
    
End Sub

Private Sub Form_Resize()

    With m_rShowAllResults
        .Top = Me.ScaleHeight - 18
        .Left = 28
        .Right = Me.ScaleWidth
        .Bottom = .Top + 30
    End With

    With m_paSearchIconDestinationPosition
        .Y = Me.ScaleHeight - 18
        .X = 7
    End With

    m_TextDisplayWidth = Me.ScaleWidth - m_cIconSize - m_cBorderSize

    mDx.SetDimensionVars
    
    SetupScrollbars
    ReCalculateVisibleNodes

End Sub

Private Sub SetupScrollbars()

    If scrVerticle.Visible Then
        scrVerticle.Move Me.ScaleWidth - scrVerticle.Width, 0, scrVerticle.Width, Me.ScaleHeight
        
        mLngLeftStart = 0
        mvRolloverPos.X = 0
        
        m_TextDisplayWidth = Me.ScaleWidth - m_cIconSize - m_cBorderSize
    Else
        mLngLeftStart = M_CBORDERSIZE_NOSCROLL
        mvRolloverPos.X = M_CROLLOVER_LEFTSTART_NOSCROLL
        
        m_TextDisplayWidth = Me.ScaleWidth - m_cIconSize
    End If
        
    scrVerticle.LargeChange = Me.ScaleHeight - (m_cNodeSpace * 2)
    mlngLowerBound = Me.ScaleHeight - 17

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngClickIndex As Long

    If m_nTargetNode Is Nothing Then
        Exit Sub
    End If

    If Button = vbLeftButton Then
        ActionTargetNode
    ElseIf Button = vbRightButton Then
        ShowContextMenu
    End If

End Sub

Private Function ActionTargetNode() As Boolean

    On Error GoTo Trapper

    ActionTargetNode = True

    If (m_nTargetNode Is Nothing) Then
        ActionTargetNode = False
        Exit Function
    End If
    
    If Left$(m_nTargetNode.Tag, 2) = "//" Then
        
        If m_nTargetNode.Tag = "//SYS_SHOW_ALL" Then
            If m_bRestrictSearch = False Then
                m_bRestrictSearch = True
            Else
                m_bRestrictSearch = False
                m_bExceedsLimits = False
            End If

            bKeyWordChanged = True
            Me.ForceRepaint
        End If
    
    ElseIf Not m_nTargetNode.IsFile Then

        If m_nTargetNode.Expanded Then
            m_nTargetNode.Expanded = False
                
            mTreeViewData.ColapseChildren m_nTargetNode
        Else
            m_nTargetNode.Expanded = True
        End If
                
        Form_Paint
        
    Else
        RaiseEvent onClick(m_nTargetNode)
    End If
    
    Exit Function
Trapper:
    Logger.Error Err.Description, "ActionTargetNode"

End Function

Public Sub ForceRepaint()
    Form_Paint
End Sub

Private Sub Form_Paint()

Dim recRolloverPos As POINTL
Dim recSize As RECT

Dim lngHorizontalMax As Long
Dim lngVerticleMax As Long

Dim lngClipHorizontal As Long
Dim lngClipVerticle As Long
    
Dim rVisibleArea As RECT
Dim rSrcAfterRollover As RECT

Dim recQuery As RECT
    
Dim lngNewVScrollMax As Long
Dim thisTV_Type As TV_Type

Dim lngSpareNodeSpaces As Long
Dim lngSpareNodeSpaces2 As Long
    
    Set mcVisibleNodes = New Collection
    mLngLevel = m_cBorderSize
    
    mrTextPos.Bottom = 0
    mrTextPos.Left = mLngLeftStart
    mrTextPos.Right = 0
    mrTextPos.Top = mLngTopStart
    
    SetScrollbarValues recSize, lngVerticleMax, lngClipVerticle

    rVisibleArea.Top = 0
    rVisibleArea.Left = 0
    rVisibleArea.Bottom = Me.ScaleHeight
    rVisibleArea.Right = Me.ScaleWidth - lngClipVerticle
    
    m_seperator_index = 0

    ResetSeperatorPositions
    
    mDx.OpenScene m_backBrush
    
        If Not (m_nTargetNode Is Nothing) Then
            If m_bKeyboardMode = True Then
            
                mDx.AddSpriteEX m_dcRollover.hdc, mvRolloverPos, mvRolloverDesc, m_dcRollover.Width, m_dcRollover.Height
                
                mvRolloverIconPos.Y = m_rTextRolloverPos.Top
                mvRolloverIconPos.X = m_nTargetNode.Left - m_cNodeSpace
            End If
        End If
    
        If Len(mstrKeyWord) > 0 Then
            If bKeyWordChanged Then
                bKeyWordChanged = False
            
                mTreeViewSearchResults.ClearNodes
                Set mTreeViewSearchResults.RootNode.Children = New Collection
    
                m_bSearchMode = True

                m_rrolloverSize.Bottom = 0
                m_rrolloverSize.Right = 0

                If m_bRestrictSearch Then
                    m_bExceedsLimits = False
                    
                    
                    For Each thisTV_Type In m_colTypes
                                                
                                                If thisTV_Type.Children.count > 0 Then
                                                        
                                                        m_lastTypeCount = thisTV_Type.Children.count
                                                
                                                        'Debug.Print "Results of: " & thisTV_Type.Caption & " @ " & mstrKeyWord & " @ " & thisTV_Type.Children.count
                                                
                                                        mTreeViewSearchResults.RootNode.copyNode thisTV_Type.Node
                                                        
                                                        'm_seperator_max = 0
                                                        
                                                        If thisTV_Type.AllowQuery Then
                                                                recQuery = mTreeViewData.QueryCollection(mstrKeyWord, thisTV_Type.Children, m_nodeDisplayLimit, m_bExceedsLimits, lngSpareNodeSpaces)
                                                                m_programCount = recQuery.Bottom / m_cNodeSpace
                                                        
                                                        Else
                                                                recQuery = mTreeViewData.ShowAll(thisTV_Type.Children, m_nodeDisplayLimit, m_bExceedsLimits)
                                                        End If
                                                        
                                                        lngSpareNodeSpaces2 = lngSpareNodeSpaces
                                                        
                                                        m_rrolloverSize.Bottom = m_rrolloverSize.Bottom + recQuery.Bottom + (m_cNodeSpace + M_SEPARATOR_GAP)
                                                        
                                                        If m_rrolloverSize.Right < recQuery.Right Then
                                                                m_rrolloverSize.Right = recQuery.Right
                                                        End If
                                                
                                                Else
                                                        m_bSearchResultsEmpty = False
                                                
                                                End If
                                        Next
                                        
                                        
                Else
                                
                    For Each thisTV_Type In m_colTypes
                                                If thisTV_Type.Children.count > 0 Then
                                                        m_lastTypeCount = thisTV_Type.Children.count
                                                
                                                        'Debug.Print "Results of: " & thisTV_Type.Caption
                                                
                                                        mTreeViewSearchResults.RootNode.copyNode thisTV_Type.Node
                                                        
                                                        'm_seperator_max = 0
                                                        
                                                        If thisTV_Type.AllowQuery Then
                                                                recQuery = mTreeViewData.QueryCollection(mstrKeyWord, thisTV_Type.Children)
                                                        Else
                                                                recQuery = mTreeViewData.ShowAll(thisTV_Type.Children)
                                                        End If
                                                        
                                                        m_rrolloverSize.Bottom = m_rrolloverSize.Bottom + recQuery.Bottom + (m_cNodeSpace + M_SEPARATOR_GAP)
                                                        
                                                        If m_rrolloverSize.Right < recQuery.Right Then
                                                                m_rrolloverSize.Right = recQuery.Right
                                                        End If
                                                Else
                                                        m_bSearchResultsEmpty = False
                                                        
                                                End If
                    Next
                End If

                m_bSearchMode = False
                    
                SetScrollbarValues recSize, lngVerticleMax, lngClipVerticle
            End If
            
            mTreeViewSearchResults.irriterateNode mTreeViewSearchResults.RootNode
            
            If m_bExceedsLimits Then
                'Debug.Print "Print Test!"
                
                mcVisibleNodes.Add m_nShowAllResults
                'mcVisibleNodes.Add m_nShowAllResults
                
                mDx.AddSpriteEX m_dcSearchIcon.hdc, m_paSearchIconDestinationPosition, m_paSearchIconSourcePosition, 16, 16
                mDx.DrawText m_nShowAllResults.Caption, m_rShowAllResults
            End If
        Else
            m_bSearchMode = False
            m_bSearchResultsEmpty = False
            
            
            mTreeViewData.irriterateNode mTreeViewData.RootNode

        End If

        
        If scrVerticle.Max <> lngVerticleMax Then
            lngNewVScrollMax = lngVerticleMax + m_cBorderSize
            
            If lngNewVScrollMax < 16038 Then
                scrVerticle.Max = lngVerticleMax + m_cBorderSize
            Else
                scrVerticle.Max = 16038
            End If
        End If
    
        'Debug.Print "scrVerticle_Max:: " & scrVerticle.Max
    
    mDx.PresentSceneEx rVisibleArea
End Sub

Private Sub SetScrollbarValues(ByRef recSize As RECT, ByRef lngVerticleMax As Long, ByRef lngClipVerticle As Long)

    If Len(mstrKeyWord) > 0 Then
        recSize = m_rrolloverSize
    Else
        recSize = mTreeViewData.Size
    End If

    lngVerticleMax = (recSize.Bottom - Me.ScaleHeight)
    
    If lngVerticleMax < 0 Then
        If scrVerticle.Visible Then
            scrVerticle.Visible = False
        End If
    Else
        scrVerticle.Visible = True
        
        lngClipVerticle = scrVerticle.Width
    End If
    
    SetupScrollbars
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mSearchProvider
    
    UnhookWindow Me.hWnd

    Set mTreeViewData = Nothing
    Set m_dcRollover = Nothing

End Sub

Sub SelectVisibleItem(ByRef lngItemIndex As Long)

Dim lngActualNode As Long
    lngActualNode = RoundIt((scrVerticle.Value + 10) / m_cNodeSpace, 1) + lngItemIndex

    If lngActualNode > mcVisibleNodes.count Then
        lngActualNode = mcVisibleNodes.count
    End If

    m_bKeyboardMode = True

    If m_bSearchResultsEmpty Then
        Exit Sub
    End If
    
    m_lngNodeIndex = lngActualNode
    mLngRounded = m_cNodeSpace * lngActualNode

    UpdateRolloverPosition
    MoveRollover

End Sub

Sub SelectFirstVisibleItem()

Dim lngFirstVisibleNode As Long
    lngFirstVisibleNode = RoundIt((scrVerticle.Value + 10) / m_cNodeSpace, 1)

    m_pCursor.X = -1
    m_bKeyboardMode = True

    If m_bSearchResultsEmpty Then
        Exit Sub
    End If
    
    m_lngNodeIndex = lngFirstVisibleNode - 1
    mLngRounded = m_cNodeSpace * m_lngNodeIndex
    
    moveRolloverDown

    UpdateRolloverPosition
    MoveRollover

End Sub

Sub SelectFirstItem()

Dim bAbort As Boolean

    m_pCursor.X = -1

    m_bKeyboardMode = True

    If m_bSearchResultsEmpty Then
        Exit Sub
    End If
    
    m_lngNodeIndex = 0
    mLngRounded = m_cNodeSpace * m_lngNodeIndex

    bAbort = False
    
    While bAbort = False
        If (m_lngNodeIndex = mcVisibleNodes.count) Then
            bAbort = True
        Else
            mLngRounded = mLngRounded + m_cNodeSpace
            m_lngNodeIndex = m_lngNodeIndex + 1
            
            If Not mcVisibleNodes(m_lngNodeIndex).IconPosition = MC_SEPERATOR_NODE Then
                bAbort = True
            End If
        End If
    Wend
    
    mvRolloverPos.Y = CalculateRolloverYFromCurrentIndex
    UpdateRolloverPosition

End Sub

Sub SelectLastItem()
    
    On Error GoTo Trapper
    
    m_bKeyboardMode = True
    
    m_pCursor.X = -1

    If m_bSearchResultsEmpty Then
        Exit Sub
    End If
    
    If scrVerticle.Visible Then
        scrVerticle.Value = scrVerticle.Max
    End If
    
    m_lngNodeIndex = mcVisibleNodes.count
    mLngRounded = (m_cNodeSpace * m_lngNodeIndex)
    
    If mcVisibleNodes(m_lngNodeIndex).IconPosition = MC_SEPERATOR_NODE Then
        moveRolloverUp
    End If

    UpdateRolloverPosition
    MoveRollover
    Form_Paint

    Exit Sub
Trapper:
    Logger.Error Err.Description, "SelectLastItem"

End Sub

Sub moveRolloverUp()

    On Error GoTo Trapper

    If m_lngNodeIndex < 0 Then
        'Can't go up any further
        Me.SelectLastItem
        Exit Sub
    End If

Dim bAbort As Boolean
    bAbort = False
            
    While bAbort = False
        If m_lngNodeIndex = 1 Then
            bAbort = True
            
            m_bKeyboardMode = False
            RaiseEvent onNotifyAllPrograms

            Form_Paint
            
        Else
            mLngRounded = mLngRounded - m_cNodeSpace
            m_lngNodeIndex = m_lngNodeIndex - 1
            
            If Not mcVisibleNodes(m_lngNodeIndex).IconPosition = MC_SEPERATOR_NODE Then
                bAbort = True
            End If
        End If
    Wend
    
    mvRolloverPos.Y = CalculateRolloverYFromCurrentIndex()
    UpdateRolloverPosition
    
    Exit Sub
Trapper:
    Logger.Error Err.Description, "moveRolloverUp"
        
End Sub

Sub moveRolloverDown()
    On Error GoTo Trapper
    
    If m_lngNodeIndex < 0 Then
        'Can't go up any further
        Me.SelectFirstVisibleItem
        Exit Sub
    End If

Dim bAbort As Boolean
    bAbort = False
    
    While bAbort = False
        If (m_lngNodeIndex = mcVisibleNodes.count) Then
            bAbort = True
            
            m_bKeyboardMode = False
            RaiseEvent onNotifyAllPrograms
            
            Form_Paint
        Else
            mLngRounded = mLngRounded + m_cNodeSpace
            m_lngNodeIndex = m_lngNodeIndex + 1
            
            If Not mcVisibleNodes(m_lngNodeIndex).IconPosition = MC_SEPERATOR_NODE Then
                bAbort = True
            End If
        End If
    Wend
    
    mvRolloverPos.Y = CalculateRolloverYFromCurrentIndex()
    UpdateRolloverPosition
    
    Exit Sub
Trapper:
    Logger.Error Err.Description, "moveRolloverDown"

End Sub

Private Function IHookSink_WindowProc(hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long

    On Error GoTo Handler

    'form-specific handler
     Select Case uMsg
     
        Case WM_MOUSELEAVE
            m_trackingMouse = False
            
            m_bKeyboardMode = False
            m_lngNodeIndex = -1
            
            If m_contextMenu Is Nothing Then Form_Paint
     
        Case WM_MOUSEMOVE
            'Prove the Mouse changed position
            If lParam <> lngLastPosition Then
                lngLastPosition = lParam
            
                If GetActiveWindow <> Me.hWnd Then
                    'Set KeyBoard.mFocusObj = Me.g_iKeyboard
                    'Me.SetFocus
                End If
            End If
     
        Case WM_MOUSEWHEEL
            With scrVerticle
                If .Visible = True Then
                    If wParam > -1 Then
                        'Up
                        If (.Value - SCROLL_CHANGE) > .Min Then
                            .Value = .Value - SCROLL_CHANGE
                        Else
                            .Value = .Min
                        End If
                    Else
                        'Down
                        If (.Value + SCROLL_CHANGE) < .Max Then
                            .Value = .Value + SCROLL_CHANGE
                        Else
                            .Value = .Max
                        End If
                    End If
                End If
            
            End With
        
     End Select

Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       CallOldWindowProcessor(hWnd, uMsg, wParam, lParam)

End Function

Private Sub m_contextMenu_onClick(theItemTag As String)
    On Error GoTo Handler
        
Dim theCommand As String
Dim thisProgram As clsProgram

    m_contextMenu_onInActive
    theCommand = theItemTag
    
    If m_nLastSelectedNode Is Nothing Then
        Logger.Error "m_nLastSelectedNode was unexpectedly empty!", "m_contextMenu_onClick"
        Exit Sub
    End If
    
    Select Case UCase$(theCommand)
    
    Case "COPYTOVIPAD"
        ShellEx m_addToViPadCommand
    
    Case "COPYTODESKTOP"
        If FSO.FolderExists(sVar_Reg_Desktop) And Not FSO.FileExists(sVar_Reg_Desktop & "\" & GetFileName(m_nLastSelectedNode.Tag)) Then
            FSO.CopyFile m_nLastSelectedNode.Tag, sVar_Reg_Desktop & "\" & GetFileName(m_nLastSelectedNode.Tag), False
        End If
    
    Case "EXPAND", "COLLAPSE", "OPEN"
        ActionTargetNode
    
        Case "EXPLORE"
                ExplorerRun m_nLastSelectedNode.Tag
        
    Case "RUNASADMIN"
        RaiseEvent onRequestCloseStartMenu
        'ShellEx m_nLastSelectedNode.Tag, "runas"
        ShellExecuteW 0, StrPtr("runas"), StrPtr(m_nLastSelectedNode.Tag), 0, StrPtr(""), SW_SHOWNORMAL
        
        thisProgram.IncreaseCount
        RaiseEvent onRequestRecentProgramsRefresh
    
    Case "TOGGLEPIN"
        Settings.Programs.TogglePin_ElseAddToPin_ByProgram CreateProgramFromNode(m_nLastSelectedNode)
        'RaiseEvent onRequestRecentProgramsRefresh
        
    Case "PROPERTIES"
        RaiseEvent onRequestCloseStartMenu
        ShellEx m_nLastSelectedNode.Tag, "properties"
        
    End Select
    
    Exit Sub
Handler:
    Logger.Error Err.Description, "m_contextMenu_onClick"
End Sub

Private Sub m_contextMenu_onInActive()
    If Not m_contextMenu Is Nothing Then
        Unload m_contextMenu
        Set m_contextMenu = Nothing
    End If
    
    Form_Paint
End Sub

Private Sub mSearchProvider_onNewItem()

Dim totalItemsCount As Long
Dim thisType As TV_Type

    For Each thisType In m_colTypes
        totalItemsCount = totalItemsCount + thisType.Children.count
    Next

    If totalItemsCount >= DISPLAY_MAX Then
        mSearchProvider.SendAbortSignal
    End If

    bKeyWordChanged = True 'This forces the paint routine to requery the items in the collection
    ForceRepaint
End Sub

Private Sub mSearchProvider_onNewService(theSearchObject As CustomSearchSlave)

Dim m_TVTType As New TV_Type
Dim desiredPosition As Long

    If UCase$(theSearchObject.SearchType) = "APPS" Then
        desiredPosition = 2
    End If
    
    m_TVTType.Caption = GetPublicString("strFiles")
    m_TVTType.DisplayLimit = 5
    m_TVTType.AllowQuery = False
    
    Set m_TVTType.Children = theSearchObject.Results
    
    AddType m_TVTType, desiredPosition
End Sub

Private Sub mSearchProvider_onUpdateCollection(theSearchObject As CustomSearchSlave)

Dim thisNode As TV_Type

    If Not ExistInCol(m_colTypes, theSearchObject.SearchType) Then
        Exit Sub
    End If

    Set thisNode = m_colTypes(theSearchObject.SearchType)
    Set thisNode.Children = theSearchObject.Results
End Sub

Private Sub mTreeViewData_onCalcNodeWidth(strCaption As String, lngWidth As Long)
    'lngWidth = Me.CalcNodeWidth(strCaption)
End Sub

Private Sub mTreeViewData_onUpLevel()
    mLngLevel = mLngLevel + m_cNodeSpace
End Sub

Private Sub mTreeViewData_onDownLevel()
    mLngLevel = mLngLevel - m_cNodeSpace
End Sub

Private Sub mTreeViewData_onNode(targetNode As INode)
    
'Dim seperatorPosition As POINTL
Dim selectedSprite As pcMemDC
Dim thisText As String
Dim thisColType As TV_Type
Dim typeChildCount As Long

    If m_bSearchMode Then

        mTreeViewSearchResults.RootNode.copyNode targetNode
        
    Else
    
        Dim lngNewX As Long
    
        lngNewX = mLngLeftStart + mLngLevel
        
        mcVisibleNodes.Add targetNode
        
        mvIconIndex.Y = targetNode.IconPosition * m_cIconSize
        mvIconPos.X = lngNewX
        
        'Check for seperator
        If targetNode.IconPosition < 0 Then
         
            If targetNode.IconPosition = MC_SEPERATOR_NODE Then
            

                Set thisColType = m_colTypes(targetNode.Caption)
                
                If Not thisColType.AllowQuery Then
                    typeChildCount = m_colTypes(targetNode.Caption).Children.count
                Else
                    typeChildCount = m_programCount
                End If
                    
                If typeChildCount = 0 Then
                    mcVisibleNodes.Remove mcVisibleNodes.count
                    Exit Sub
                End If
            
                mrTextPos.Left = lngNewX
                
                mvIconPos.Y = (mrTextPos.Top - 2)
                mvIconPos.X = M_CBORDERSIZE_NOSCROLL
                
                mDx.Font = m_separatorFont
                mDx.TextColour = m_Seperator_fontColour
                
                If ExistInCol(m_dcLib, LCase$(targetNode.Caption)) Then
                    Set selectedSprite = m_dcLib(LCase$(targetNode.Caption))
                End If
                
                If selectedSprite Is Nothing Then
                    If Not thisColType.AllowQuery Then
                        thisText = targetNode.Caption & " (" & typeChildCount & ")"
                    Else
                        thisText = thisColType.Caption & " (" & m_programCount & ")"
                    End If
                
                    mDx.DrawText thisText, CreateRect(5, mrTextPos.Top, Me.ScaleWidth, mrTextPos.Top + 30), DT_LEFT
                    mDx.AddSpriteEX m_dcTreeSeparator.hdc, CreatePointL(mrTextPos.Top + 11, mDx.GetTextRect(thisText).Right + 10), m_paBlank, m_dcTreeSeparator.Width, 25
                Else
                    mDx.AddSpriteEX selectedSprite.hdc, CreatePointL(mrTextPos.Top + 2, 2), m_paBlank, selectedSprite.Width, 25
                End If
                
                mDx.TextColour = m_originalTextColour
                mDx.Font = mStdFont
    
                'SCRUB
                mrTextPos.Top = mrTextPos.Top + m_cNodeSpace + M_SEPARATOR_GAP
                
                If m_seperator_index <= UBound(m_seperator) And _
                    m_seperator_index >= LBound(m_seperator) Then
                    
                    m_seperator(m_seperator_index) = mvIconPos.Y
                    m_seperator_inuse(m_seperator_index) = True
                    
                    m_seperator_index = m_seperator_index + 1
                    'm_seperator_max = m_seperator_max + 1
                End If
                
            Else
            

                mrTextPos.Left = lngNewX + m_cNodeSpace
                
                mvIconPos.Y = mrTextPos.Top
                mvIconPos.X = lngNewX
    
                mrTextPos.Right = m_TextDisplayWidth
                mrTextPos.Bottom = mrTextPos.Top + m_cNodeSpace
                
                If mrTextPos.Top > -m_cNodeSpace Then
                    mDx.DrawText targetNode.Caption, mrTextPos, DT_LEFT Or DT_NOPREFIX Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
                    targetNode.Left = mrTextPos.Left
                    
                    mvIconIndex.X = 0
                    mvIconIndex.Y = 0
                    
                    'Only draw what we can see
                    If mvIconPos.Y < (mlngLowerBound + m_cNodeSpace) Then
                    
                        If targetNode.IconPosition = -3 Then

                            'Set targetNode.Icon = New ViIcon
                            'targetNode.Icon.LoadIconFromFile targetNode.Tag
                            
                            'always set to loaded
                            'targetNode.IconPosition = -4
                            
                            Dim newIcon
                            Set newIcon = IconManager.GetViIcon(targetNode.Tag)
                            
                            Set targetNode.Icon = IconManager.GetViIcon(targetNode.Tag)
                            
                            If targetNode.Icon Is Nothing Then
                                Set targetNode.Icon = New ViIcon
                            End If
                            targetNode.IconPosition = -4
                        End If
                    
                        If targetNode.IconPosition = -4 Then

                            If Not targetNode.Icon Is Nothing Then
                                targetNode.Icon.DrawIcon mDx.hdc, mvIconPos.X, mvIconPos.Y
                            Else
                                Logger.Error "invalid icon specified:: " & targetNode.Tag, "mTreeViewData_onNode"
                            End If
                            'ExtractIconEx targetNode.Tag, mDx.hdc, 16, mvIconPos.X, mvIconPos.Y
                        Else
                            mDx.AddSpriteEX m_dcIndexIcons.hdc, mvIconPos, mvIconIndex, m_cIconSize, m_cIconSize
                        End If
                    End If
                End If
                
                mrTextPos.Top = mrTextPos.Top + m_cNodeSpace
            End If
        Else
            mrTextPos.Left = lngNewX + m_cNodeSpace
            
            mvIconPos.Y = mrTextPos.Top
            mvIconPos.X = lngNewX

            mrTextPos.Right = m_TextDisplayWidth
            mrTextPos.Bottom = mrTextPos.Top + m_cNodeSpace
            
            'Only draw what we can see
            If mrTextPos.Top > mlngLowerBound Then
                mTreeViewData.Abort
                mTreeViewSearchResults.Abort
            End If
            
            If mrTextPos.Top > -m_cNodeSpace Then
                mDx.DrawText targetNode.Caption, mrTextPos, DT_LEFT Or DT_NOPREFIX Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
                targetNode.Left = mrTextPos.Left
            End If
            
            mrTextPos.Top = mrTextPos.Top + m_cNodeSpace
        End If
    
    End If

End Sub

Private Sub picRollover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub mTreeViewSearchResults_onDownLevel()
    mLngLevel = mLngLevel - m_cNodeSpace
End Sub

Private Sub mTreeViewSearchResults_onNode(targetNode As INode)
    mTreeViewData_onNode targetNode
End Sub

Private Sub mTreeViewSearchResults_onUpLevel()
    mLngLevel = mLngLevel + m_cNodeSpace
End Sub

Private Sub scrVerticle_Change()
    m_lngNodeIndex = FindNodeIndex(m_pCursor.Y)
    mvRolloverPos.Y = CalculateRolloverYFromCurrentIndex()
    
    UpdateRolloverPosition

    mLngTopStart = -scrVerticle.Value + m_cBorderSize

    MoveRollover
    Form_Paint
End Sub

Private Sub scrVerticle_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    KeyCode = 0
End Sub

Private Sub scrVerticle_Scroll()
    scrVerticle_Change
End Sub



