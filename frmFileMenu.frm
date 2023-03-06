VERSION 5.00
Begin VB.Form frmFileMenu 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "Vista_RecentItems"
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   51
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   Begin VB.Timer timIconDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   360
   End
End
Attribute VB_Name = "frmFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_toolTip As ViToolTip
Private WithEvents m_contextMenu As frmVistaMenu
Attribute m_contextMenu.VB_VarHelpID = -1

Private m_hdcBackBuffer As ISoftX

Private m_gobjBack As Long
Private m_pBrush As Long

Private m_colLabels As Collection
Private m_lngRolloverIndex As Long

Private m_hdcRollover As pcMemDC

Private m_IconSheet As pcMemDC
Private m_IconSheetOver As pcMemDC

Private m_paCursorPos As POINTL
Private m_lngFormWidth As Long
Private m_trackingMouse As Boolean

Private m_paRolloverPosition As POINTL
Private m_Brush As GDIBrush
Private m_selectedItem As INode

Private Const C_FILEHEIGHT As Long = 22
'---
Private Const C_HEIGHT As Integer = 22

Public Event onClickItem(strPath As String)
Public Event onInActive()

Implements IHookSink

Private m_logger As SeverityLogger

Property Get Logger()
    Set Logger = m_logger
End Property

Function ShowContextMenu() As Boolean
    If (m_selectedItem Is Nothing) Then
        ShowContextMenu = False
        Exit Function
    End If
    
    If Not m_contextMenu Is Nothing Then Unload m_contextMenu
    Set m_contextMenu = BuildGenericFileContextMenu(m_selectedItem.Tag)
    
    Logger.Trace "Attemping to show context menu", "ShowContextMenu"
    m_contextMenu.Resurrect True, Me
End Function

Property Let SelectItem(ByVal lngNewItemIndex As Long)
    On Error GoTo Handler
    
    If lngNewItemIndex = m_lngRolloverIndex Then
        'Already selected
        Exit Property
    End If
    
    If ExistInCol(m_colLabels, (lngNewItemIndex + 1)) Then
        Set m_selectedItem = m_colLabels(lngNewItemIndex + 1)
    
        m_lngRolloverIndex = lngNewItemIndex
        m_paRolloverPosition.Y = m_lngRolloverIndex * C_FILEHEIGHT
        
        m_toolTip.Hide
        m_toolTip.SetToolTip m_selectedItem.Tag
        
        Form_Paint
    End If
    
Handler:
    If Err Then
        MsgBox Err.Description
    End If
End Property

Private Function MakeIconSheet() As Boolean
    On Error GoTo Handler

Dim iconCount As Long
Dim iconIndex As Long
Dim Y As Long

    iconCount = m_colLabels.count
    
    For iconIndex = 1 To iconCount
        ExtractIconEx m_colLabels(iconIndex).Tag, m_IconSheet.hdc, 16, 0, Y
        
        BitBlt m_IconSheetOver.hdc, 0, Y, 16, 16, m_hdcRollover.hdc, 1, 3, vbSrcCopy
        ExtractIconEx m_colLabels(iconIndex).Tag, m_IconSheetOver.hdc, 16, 0, Y
        
        Y = Y + C_HEIGHT
    Next

    MakeIconSheet = True
    Exit Function
Handler:
End Function

Private Function GetItems_MRU()

Dim sMRUListEx As String
Dim lngRegIndex As Double
Dim sPath As String

Dim sMRU As String
Dim sSeekFile As String
Dim sSeekPath As String

Dim cItems As New Collection
Dim sP() As String
Dim recentDocsRegKey As RegistryKey

    Set Me.Picture = Nothing
    Set GetItems_MRU = cItems
    
    m_paRolloverPosition.Y = -m_hdcRollover.Height
    
    sPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
    Set recentDocsRegKey = Registry.CurrentUser.OpenSubKey(sPath)
    
    If recentDocsRegKey Is Nothing Then
        Logger.Error "Could not open registry key", "GetItems_MRU"
        Exit Function
    End If
    
    sMRUListEx = recentDocsRegKey.GetValue("MRUListEx")

    While (Len(sMRUListEx) > 4) And (cItems.count < 15)
        lngRegIndex = GetDWord(ExtractBytes(sMRUListEx, 4))
        
        sMRU = recentDocsRegKey.GetValue(lngRegIndex)
        
                                '14 00    00 00
        GetStringByString sMRU, Chr$(&H14) & Chr$(0)
                                '00 00    00
        sSeekFile = GetStringByString(sMRU, Chr$(0) & ChrB$(0))
        
        If Not g_WindowsXP Then
            sSeekPath = sVar_Reg_StartMenu_CurrentUserRecentItems & "\" & sSeekFile & ".lnk"
        Else
            sSeekPath = sVar_Reg_StartMenu_CurrentUserRecentItems & "\" & sSeekFile
        End If
            
        If Not isFolderShortcut(sSeekPath) Then
            cItems.Add ExtOrNot(sSeekFile) & "*" & sSeekPath
        End If
    Wend


End Function

Private Function GetItems(szPath As String)

    On Error GoTo Handler

Dim recentFiles As Scripting.Folder
Dim thisFile As Scripting.File
Dim itemCollection As Collection
Dim folderPath As String

    Set itemCollection = New Collection
    Set GetItems = itemCollection

    Set Me.Picture = Nothing
    Set m_colLabels = New Collection
    folderPath = szPath
    
    If FSO.FolderExists(folderPath) Then
        Set recentFiles = FSO.GetFolder(folderPath)
        
        For Each thisFile In recentFiles.Files
        
            If FSO.FileExists(ResolveLink(thisFile.Path)) Then
                itemCollection.Add ExtOrNot(thisFile.Name) & "*" & thisFile.Path
                
                If itemCollection.count = 15 Then
                    Exit For
                End If
            End If
        Next
    End If

Handler:
End Function

Function PopulateFromPath(szPath As String)

Dim lngRegIndex As Double

Dim sMRU As String
Dim sSeekFile As String
Dim sSeekPath As String

Dim cItems As New Collection
Dim sP() As String

Dim lngTextWidth As Long
Dim lngItemIndex As Long
Dim theFileName As String

    Set Me.Picture = Nothing
    
    ReInitialize
    
    m_paRolloverPosition.Y = -m_hdcRollover.Height
    
    If FSO.FolderExists(szPath) = False And g_WindowsXP Then
        Set cItems = GetItems_MRU
    Else
        Set cItems = GetItems(szPath)
    End If
    
    If cItems.count = 0 Then
        cItems.Add " (Empty) * \"
    End If
    
    Set cItems = StartBarSupport.QuickSortNamesAscending(cItems, vbTextCompare)
    m_lngFormWidth = 0

    For lngItemIndex = 1 To cItems.count
        sP = Split(cItems(lngItemIndex), "*")
        theFileName = sP(0)
        
        lngTextWidth = 32 + m_hdcBackBuffer.GetTextRect(theFileName).Right

        If m_lngFormWidth < 360 Then
        
            If (lngTextWidth > m_lngFormWidth) Then
                If (lngTextWidth > 360) Then
                    lngTextWidth = 360
                End If
                
                m_lngFormWidth = lngTextWidth
            End If
        End If
    Next

    Me.Width = m_lngFormWidth * Screen.TwipsPerPixelX
    
    m_IconSheet.Height = (cItems.count * C_HEIGHT) - 4
    m_IconSheet.Width = 16
    
    m_IconSheetOver.Height = m_IconSheet.Height
    m_IconSheetOver.Width = 16
    
    m_Brush.Constructor &HEAEAEA
    FillRect m_IconSheet.hdc, CreateRect(0, 0, 16, m_IconSheet.Height), m_Brush.Value
    
    'm_Brush.Constructor &HEAEAEA
    'FillRect m_IconSheetOver.hdc, CreateRect(0, 0, 16, m_IconSheetOver.Height), m_Brush.Value
    
    While cItems.count > 0
        sP = Split(cItems(1), "*")
        
        AddToVisibleItems sP(0), sP(1)

        
        cItems.Remove 1
    Wend
    
    Me.Height = (m_colLabels.count * C_HEIGHT) * Screen.TwipsPerPixelY
    Form_Paint
    timIconDelay.Enabled = True
    
    Logger.Trace "Final width of menu", "PopulateFromPath", Me.ScaleWidth
    
    Exit Function
Handler:
    'MsgBox "I tried to get file: " & sPart1 & " but something went wrong.", vbCritical

End Function

Private Sub Form_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)
    
    Set m_toolTip = New ViToolTip
    m_toolTip.AttachWindow Me.hWnd
    
    HookWindow Me.hWnd, Me
End Sub

Private Sub ReInitialize()

    Set m_colLabels = New Collection
    Set m_Brush = New GDIBrush
    Set m_hdcRollover = New pcMemDC
    Set m_IconSheet = New pcMemDC
    Set m_IconSheetOver = New pcMemDC
    
    Set m_hdcBackBuffer = New ISoftX
    
    Me.Font.Name = OptionsHelper.PrimaryFont
    
    m_hdcBackBuffer.hWnd = Me.hWnd
    m_hdcBackBuffer.Font = Me.Font
    
    m_hdcRollover.CreateFromPicture GetResourceBitmap(101)
    m_lngRolloverIndex = -1

End Sub

Private Function AreObjectsInstantiated() As Boolean

    If m_colLabels Is Nothing Or _
       m_Brush Is Nothing Or _
       m_hdcRollover Is Nothing Or _
       m_IconSheet Is Nothing Or _
       m_IconSheetOver Is Nothing Or _
       m_hdcBackBuffer Is Nothing Then

        Exit Function
    End If
    
    AreObjectsInstantiated = True
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_trackingMouse = False Then
        m_trackingMouse = TrackMouse(Me.hWnd)
    End If

    If Not m_contextMenu Is Nothing Then Exit Sub

Dim lngNewItemIndex As Long
    lngNewItemIndex = RoundIt(Y - 11, C_FILEHEIGHT) / C_FILEHEIGHT
    
    Me.SelectItem = lngNewItemIndex

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Handler
    
    If Button = vbLeftButton Then
        RaiseEvent onClickItem(m_selectedItem.Tag)
        
    ElseIf Button = vbRightButton Then
        ShowContextMenu
    End If
    
    Exit Sub
Handler:

End Sub

Private Sub Form_Paint()

Dim thisItem As INode
Dim recItemPlacement As RECT

Dim lngIndex As Long

    m_hdcBackBuffer.OpenScene
    
        
    
        m_hdcBackBuffer.SetPenColor &HB9B5B7
        m_hdcBackBuffer.DrawRectangle 0, 0, Me.ScaleWidth, Me.ScaleHeight

        m_hdcBackBuffer.SetBrushColor &HEAEAEA
        m_hdcBackBuffer.DrawRectangle 0, 0, 27, Me.ScaleHeight
        
        m_hdcBackBuffer.SetBrushColor &HF4F4F4: m_hdcBackBuffer.SetPenColor &HE0E0DF
        m_hdcBackBuffer.DrawRectangle 26, 1, 27, Me.ScaleHeight - 1
    
        m_hdcBackBuffer.SetPenColor &HFFFFFF
        m_hdcBackBuffer.DrawRectangle 27, 1, 28, Me.ScaleHeight - 1
    
        m_hdcBackBuffer.SetBrushColor &HF4F4F4

        m_hdcBackBuffer.AddSprite m_IconSheet, CreatePointL(2, 6)
        
        m_hdcBackBuffer.AddSprite m_hdcRollover, m_paRolloverPosition
        m_hdcBackBuffer.AddSpriteEX m_IconSheetOver.hdc, CreatePointL(m_paRolloverPosition.Y + 2, 6), CreatePointL(m_paRolloverPosition.Y, 0), 16, 16
        
        For Each thisItem In m_colLabels
            'DisplayIcon thisItem.Tag, lngIndex * C_HEIGHT + 3
    
            With recItemPlacement
                .Top = lngIndex * C_HEIGHT + 4
                .Bottom = .Top + 15
                
                .Right = Me.ScaleWidth
                .Left = 30
            End With
            
            m_hdcBackBuffer.DrawText thisItem.Caption, recItemPlacement, _
                DT_LEFT Or DT_NOPREFIX Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
                
            lngIndex = lngIndex + 1
        Next

        
    m_hdcBackBuffer.PresentScene

End Sub

Private Sub Form_Resize()
    TopMost Me.hWnd
    m_hdcBackBuffer.SetDimensionVars
End Sub

Private Sub AddToVisibleItems(strCaption As String, Optional strShell As String)

Dim newItem As New INode
    
    With newItem
        .Caption = strCaption
        .Tag = strShell
        
        
    End With
        
    m_colLabels.Add newItem

End Sub

Private Function DisplayIcon(strPath As String, lngY As Long)
    ExtractIconEx strPath, m_hdcBackBuffer.hdc, 16, 6, lngY
End Function

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindow Me.hWnd
    
    Set m_hdcBackBuffer = Nothing
    Set m_hdcRollover = Nothing
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
    On Error GoTo Handler

    'form-specific handler
    Select Case uMsg
       
        Case WM_MOUSELEAVE
            m_trackingMouse = False
            
            If m_contextMenu Is Nothing Then
                m_lngRolloverIndex = -1
                m_paRolloverPosition.Y = -m_hdcRollover.Height
                Form_Paint
            End If
        
        Case WM_ACTIVATE
            If wParam = WA_INACTIVE Then
                'frmEvents.m_winVistaMenu_LostFocus

                If m_contextMenu Is Nothing Then
                    RaiseEvent onInActive
                Else
                    If m_contextMenu.hWnd = lParam Then
                        'MsgBox "DUDE!"
                    End If
                End If
            End If
        
        Case Else
            ' Just allow default processing for everything else.
            IHookSink_WindowProc = _
               CallOldWindowProcessor(hWnd, uMsg, wParam, lParam)
                                           
            Exit Function
        
    End Select

    Exit Function
Handler:
    Logger.Error Err.Description, "IHookSink_WindowProc", uMsg

    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       CallOldWindowProcessor(hWnd, uMsg, wParam, lParam)
End Function

Private Sub m_contextMenu_onClick(theItemTag As String)
    If m_selectedItem Is Nothing Then Exit Sub

    m_contextMenu_onInActive
    GenericFileContextMenuHandler theItemTag, m_selectedItem.Tag
End Sub

Private Sub m_contextMenu_onInActive()
    If Not m_contextMenu Is Nothing Then
        Unload m_contextMenu
        Set m_contextMenu = Nothing
    End If
    
    On Error Resume Next
    Me.SetFocus
    If Not MouseInsideWindow(Me.hWnd) Then RaiseEvent onInActive
End Sub

Private Sub timIconDelay_Timer()
    If Not MakeIconSheet Then Exit Sub
    
    Form_Paint
    timIconDelay.Enabled = False
End Sub
