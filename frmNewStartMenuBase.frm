VERSION 5.00
Begin VB.Form frmStartMenuBase 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ViStart_MenuBase"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timJumplistUpdater 
      Interval        =   2000
      Left            =   1200
      Top             =   1800
   End
   Begin VB.Timer timMorphToJumpList 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2400
      Top             =   600
   End
   Begin VB.Timer timRolloverDelay 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3120
      Top             =   960
   End
   Begin VB.Timer timAutoClick 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Timer timFade 
      Interval        =   15
      Left            =   2040
      Top             =   2280
   End
   Begin VB.Timer timTreeViewSearch 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2280
      Top             =   1320
   End
End
Attribute VB_Name = "frmStartMenuBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_OptionsGap As Single

Private m_groupMenuFontSize As Single
Private m_AllProgramsFontSize As Single
Private m_ShutDownTextFontSize As Single

Private WithEvents m_optionDialog As frmControlPanel
Attribute m_optionDialog.VB_VarHelpID = -1
Private WithEvents m_programMenu As frmTreeView
Attribute m_programMenu.VB_VarHelpID = -1
Private WithEvents m_recentPrograms As frmFreq
Attribute m_recentPrograms.VB_VarHelpID = -1
Private WithEvents m_powerMenu As frmVistaMenu
Attribute m_powerMenu.VB_VarHelpID = -1
Private WithEvents m_RecentItems As frmFileMenu
Attribute m_RecentItems.VB_VarHelpID = -1
Private WithEvents m_jumpListDrawer As JumpListDrawer
Attribute m_jumpListDrawer.VB_VarHelpID = -1
Private WithEvents m_navigationDraw As ViNavigationPane
Attribute m_navigationDraw.VB_VarHelpID = -1

Private m_toolTip As ViToolTip

Private m_StartMenuMaskBitmap As GDIBitmap
Private m_StartMenuMaskDC As GDIDC
Private m_StartMenuMaskInvertedDC As pcMemDC

Private m_Bitmap As GDIPBitmap
Private m_BitmapGraphics As GDIPGraphics
Private m_graphics As GDIPGraphics
Private m_DesktopBitmap As GDIPBitmap

Private m_RolloverToShow As String

Private m_font As GDIPFont

Private m_FontNormal As GDIFont
Private m_FontItalic As GDIFont

Private m_ShutDownBrush2 As GDIPBrush
Private m_ShutDownBrush As GDIPBrush

Private m_GroupOptionsBrush As GDIPBrush
Private m_AllProgramsBrush As GDIPBrush

Private m_BlackBrush As GDIPBrush

Private m_Pen As GDIPPen
Private m_Path As GDIPGraphicPath

Private m_GroupOptionsFontFamily As GDIPFontFamily
Private m_AllProgramsFontFamily As GDIPFontFamily
Private m_ShutDownTextFontFamily As GDIPFontFamily

'Objects
Private m_background As GDIPImage
Private m_BackGroundPinned As GDIPImage

Private m_AllPrograms As RolloverButton
Private m_ArrowAllPrograms As RolloverButton
Private m_options() As OptionText

Private WithEvents m_searchText As VistaSearchBox
Attribute m_searchText.VB_VarHelpID = -1

Private m_shutDownButton As PowerOptionButton
Private m_logOffButton As PowerOptionButton
Private m_arrowButton As PowerOptionButton

Private m_blendFunc32bpp As BLENDFUNCTION
Private m_layeredData As LayerdWindowHandles

Private m_windowPosition As POINTL
Private m_ignoreActivation As Boolean

Private m_AllProgramsCaption As String
Private m_ShutDownCaption As String

Private m_trackingMouse As Boolean

Private m_currentFadeInImage As frmRolloverImage
Private m_currentFadeOutImage As frmRolloverImage

Private m_FadeOuts As Collection

Private m_rolloverImagesPlaceHolder As POINTF
Private m_userPicturePlaceHolder As POINTF

Private m_recentItems_Y As Long
Private m_recentItemsAutoClick As Boolean

Private m_programMenuHeight As Long

Private m_searchBox_Y As Long

Private m_searchBoxForeFontItalic As GDIFont
Private m_searchBoxNormalFont As GDIFont

Private m_initalized As Boolean

Private m_searchingStarted As Boolean
Private m_research As Boolean
Private m_jumpListMode As Boolean
Private m_reBlurcallback As Boolean
Private m_layeredMode As Boolean

Private m_theMatrix As ColorMatrix
Private m_theMatrix2 As ColorMatrix
Private m_AlphaAmount As Single

Private m_morphFrom As GDIPImage
Private m_morphTo As GDIPImage
Private m_undoLayeredOnClose As Boolean
Private m_showingMenu As Form

Private m_ShutDownTextEnabled As Boolean
Private m_JumpListEnabled As Boolean
Private m_InitialStartUp As Boolean
Private m_repaintOnce As Boolean
Private m_originalBackground As GDIPImage
Private m_contextMenu As frmVistaMenu
Private WithEvents m_userPicture As frmRolloverImage
Attribute m_userPicture.VB_VarHelpID = -1

Public BlurEnabled As Boolean

Private Type RolloverButton
    Image As GDIPImage
    Visible As Boolean
    State As Long
    AutoClick As Boolean
    
    Position As gdiplus.POINTL
    NormalButton As gdiplus.RECTL
    RolloverButton As gdiplus.RECTL
End Type

Private Type PowerOptionButton
    Image As GDIPImage
    State As Long
    AllowUpdates As Boolean
    ActionCode As Long
    AutoClick As Boolean
    Visible As Boolean
    
    Position As gdiplus.POINTL
    NormalButton As gdiplus.RECTL
    RolloverButton As gdiplus.RECTL
    PressedButton As gdiplus.RECTL
End Type

Private Type OptionText
    Caption As String
    Command As String
    RolloverPath As String
    Menu As ContextMenu
End Type

Public Event onClose()
Public Event onRequestNewResize()
Public Event onSkinChange()

Implements IHookSink

Public Function SetContextMenu(ByRef newContextMenu As frmVistaMenu)
    Set m_contextMenu = newContextMenu
End Function

Public Function InitializeCurrentSkin()

Dim r As RECTL

    CloseMe
    RaiseEvent onRequestNewResize
    
    m_initalized = False
    
    Set m_DesktopBitmap = New GDIPBitmap
    Set m_StartMenuMaskBitmap = New GDIBitmap
    Set m_StartMenuMaskDC = New GDIDC
    Set m_StartMenuMaskInvertedDC = New pcMemDC
    
    Set m_background = New GDIPImage
    Set m_BackGroundPinned = New GDIPImage
    
    Set m_shutDownButton.Image = New GDIPImage
    Set m_logOffButton.Image = New GDIPImage
    Set m_arrowButton.Image = New GDIPImage
    Set m_AllPrograms.Image = New GDIPImage
    Set m_ArrowAllPrograms.Image = New GDIPImage
    
    Set m_Path = New GDIPGraphicPath
    Set m_Pen = New GDIPPen
    Set m_GroupOptionsFontFamily = New GDIPFontFamily
        
    Set m_FontNormal = New GDIFont
    Set m_FontItalic = New GDIFont
    
    Set m_font = New GDIPFont
    Set m_GroupOptionsBrush = New GDIPBrush
    Set m_layeredData = Nothing
    
    Set Layout = New LayoutParser
    
    If Not Layout.ParseLayout(g_resourcesPath & "layout.xml") Then
        LogError "Failed to parse layout file", "StartMenuBase"
        Exit Function
    End If
    
    If BlurEnabled Then
        UnblurMe True, False
    End If
    
    If m_layeredMode Then
        UndoLayeredWindow
        m_layeredMode = False
        
    End If
    
    g_viOrb_fullHeight = Layout.ViOrb_FullHeight
    m_ShutDownTextEnabled = Not Layout.ShutDownTextSchema Is Nothing
    BlurEnabled = FileExists(g_resourcesPath & "startmenu_mask.bmp")
    m_JumpListEnabled = FileExists(g_resourcesPath & "startmenu_expanded.png")
    
    If m_JumpListEnabled Then
        If Layout.JumpListViewerSchema Is Nothing Then
            LogError "Warning:: Jumplist viewer schema is unavaliable but startmenu_expanded.png is present", "ResourcesPath"
            m_JumpListEnabled = False
        End If
    End If
    
    m_shutDownButton.AllowUpdates = True
    m_shutDownButton.Visible = Layout.ShutDownButtonSchema.Visible
    
    m_logOffButton.AllowUpdates = True
    m_logOffButton.Visible = Layout.LogOffButtonSchema.Visible
    
    m_arrowButton.AllowUpdates = True
    m_arrowButton.Visible = Layout.ArrowButtonSchema.Visible
    
    'originalBackground.FromFile g_resourcesPath & "startmenu.png"
    
    r.Left = Layout.GroupMenuSchema.Left
    r.Top = Layout.GroupMenuSchema.Top
    
    r.Bottom = r.Top + Layout.GroupMenuSchema.Height
    r.Right = r.Left + Layout.GroupMenuSchema.Width
    
    Set m_originalBackground = New GDIPImage
    m_originalBackground.FromFile g_resourcesPath & "startmenu.png"
    
    Set m_background = ReconstructBackgroundImage(m_originalBackground, r)
        
    m_shutDownButton.Image.FromFile g_resourcesPath & "bottombuttons_shutdown.png"
    m_logOffButton.Image.FromFile g_resourcesPath & "bottombuttons_logoff.png"
    m_arrowButton.Image.FromFile g_resourcesPath & "bottombuttons_arrow.png"
    m_AllPrograms.Image.FromFile g_resourcesPath & "allprograms.png"
    m_ArrowAllPrograms.Image.FromFile g_resourcesPath & "programs_arrow.png"
    
    'm_graphics.Dispose
    'm_BitmapGraphics.Dispose
    
    
    Me.Show
    Me.Width = m_background.Width * Screen.TwipsPerPixelX
    Me.Height = m_background.Height * Screen.TwipsPerPixelY
    Me.Hide
    
    With m_ArrowAllPrograms.NormalButton
        .Width = m_ArrowAllPrograms.Image.Width
        .Height = m_ArrowAllPrograms.Image.Height / 2
    End With
    
    With m_ArrowAllPrograms.RolloverButton
        .Top = m_ArrowAllPrograms.NormalButton.Height
        .Width = m_ArrowAllPrograms.Image.Width
        .Height = m_ArrowAllPrograms.Image.Height
    End With
    
    SetPowerButtonDimensions
    
    m_AllPrograms.Position.X = Layout.AllProgramsRolloverSchema.Left
    m_AllPrograms.Position.Y = Layout.AllProgramsRolloverSchema.Top
    
    m_ArrowAllPrograms.Position.X = Layout.AllProgramsArrowSchema.Left
    m_ArrowAllPrograms.Position.Y = Layout.AllProgramsArrowSchema.Top
    
    m_ArrowAllPrograms.Visible = True

    

    'If m_layeredMode Then UndoLayeredWindow
    '    SetWindowLong Me.hWnd, GWL_EXSTYLE, _
    '    GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED And Not WS_EX_TOOLWINDOW
    
    If Not BlurEnabled Then Set m_layeredData = MakeLayerdWindow(Me)
    
    If BlurEnabled Then
        SetupBlur g_resourcesPath
    Else
        m_layeredMode = True
    End If
    
    If m_JumpListEnabled Then
        SetupJumplistViewer
        m_BackGroundPinned.FromFile g_resourcesPath & "startmenu_expanded.png"
    End If
    
    SetupNavigationViewer m_originalBackground
    
    If m_ShutDownTextEnabled Then InitializeShutDownText
    
    m_GroupOptionsFontFamily.Constructor OptionsHelper.PrimaryFont
    m_font.Depreciated_Constructor OptionsHelper.PrimaryFont, CLng(m_groupMenuFontSize), FontStyleRegular
        
    m_FontNormal.Constructor OptionsHelper.PrimaryFont, 15, APIFALSE
    m_FontItalic.Constructor OptionsHelper.PrimaryFont, 15, APITRUE

    InitializeProgramMenu
    InitializeGroupMenu
    InitializeRecentPrograms
    InitializeSearchText
    InitializeAllProgramsText
    
    MakeUserRollover g_resourcesPath
    'AlignChildWindows (casues XP to crash for some reason)
    
    m_initalized = True
    
    RepaintWindow Me.hWnd
    
    ReInitSurface
    
    m_optionDialog.NavigationPanel = Settings.NavigationPane
    m_optionDialog.timReloadAbout = True
    
    RaiseEvent onSkinChange
End Function

Public Property Let Skin(ByVal szNewSkin As String)
    
    If m_initalized And LCase$(szNewSkin) = LCase$(Settings.CurrentSkin) Then
        Exit Property
    End If
    
    Settings.CurrentSkin = szNewSkin
    g_resourcesPath = sCon_AppDataPath & "_skins\" & szNewSkin & "\"

    InitializeCurrentSkin
End Property

Private Function SetupBlur(szResourcesPath As String)

    Set m_StartMenuMaskBitmap = New GDIBitmap
    Set m_StartMenuMaskDC = New GDIDC
    Set m_StartMenuMaskInvertedDC = New pcMemDC

    m_StartMenuMaskBitmap.LoadImageFromFile szResourcesPath & "startmenu_mask.bmp"
    m_StartMenuMaskDC.SelectBitmap m_StartMenuMaskBitmap
    
    m_StartMenuMaskInvertedDC.Height = m_StartMenuMaskBitmap.Height
    m_StartMenuMaskInvertedDC.Width = m_StartMenuMaskBitmap.Width
    
    m_StartMenuMaskInvertedDC.Clear
    BitBlt m_StartMenuMaskInvertedDC.hdc, 0, 0, m_StartMenuMaskBitmap.Width, m_StartMenuMaskBitmap.Height, m_StartMenuMaskDC.Handle, 0, 0, vbSrcInvert

End Function

Public Property Get SearchBoxhWnd() As Long
    SearchBoxhWnd = m_searchText.real_hWnd
End Property

Sub ActivateSearchText()
    If Not m_navigationDraw.RenameWindowHasFocus Then
        m_searchText.SetKeyboardFocus
    End If
End Sub

Sub ShowMe()
    'Checking will probably take just as long as it would to re-align them regardless
    If AlignChildWindows = False Then
        Exit Sub
    End If
    
    g_KeyboardMenuState = 0
    g_KeyboardSide = 1
    
    g_bStartMenuVisible = True
    m_ignoreActivation = False
    
    m_programMenu.Show
    If Not Settings.ShowProgramsFirst Then m_recentPrograms.Show
    ShowWindow m_searchText.hWnd, SW_SHOW
    If Not Settings.ShowProgramsFirst Then win.SetFocus m_recentPrograms.hWnd
    
    timTreeViewSearch.Enabled = True
    Me.Show
    
    SetKeyboardActiveWindow m_searchText.hWnd
    ShowAvatarRollover
    
    If m_repaintOnce Then
        m_repaintOnce = False
        
        'RepaintWindow2 Me
    End If
End Sub

Sub UpdateDesktopImage(newPosition As POINTL)
    Dim hdcSrc As Long
    Dim hDCMemory As Long
    Dim r As Long
    
    Dim theGrahipcsTemp As New GDIPGraphics
    Dim theGrahipcs As New GDIPGraphics
    
    Dim newWidth As Single
    Dim newHeight As Single
    
    'Dim encoder As New GDIPImageEncoderList
    Dim m_DesktopBitmapTemp As New GDIPBitmap
    Dim encoder As New GDIPImageEncoderList
    
    Dim desktopCopyBitmap As New GDIBitmap
    Dim desktopCopySmallDC As New GDIDC
    
    Dim desktopCopyBig As New GDIBitmap
    Dim desktopCopyBigDC As New GDIDC
    
    newWidth = Me.ScaleWidth / 4
    newHeight = Me.ScaleHeight / 4
    
    hdcSrc = GetWindowDC(GetDesktopWindow()) ' Get device context for entire
                                          ' window.

    m_DesktopBitmapTemp.CreateFromSizeFormat Me.ScaleWidth, Me.ScaleHeight, GDIPlusWrapper.Format32bppArgb
    m_DesktopBitmap.CreateFromSizeFormat newWidth, newHeight, GDIPlusWrapper.Format32bppArgb
    
    theGrahipcsTemp.FromImage m_DesktopBitmapTemp.Image
    theGrahipcs.FromImage m_DesktopBitmap.Image
    
    hDCMemory = theGrahipcsTemp.GetHDC
    
    ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, Me.ScaleWidth, Me.ScaleHeight, hdcSrc, _
            newPosition.X, newPosition.Y, vbSrcCopy)
    
    theGrahipcsTemp.ReleaseHDC hDCMemory
    ReleaseDC GetDesktopWindow(), hdcSrc

    desktopCopyBig.hBitmap = m_DesktopBitmapTemp.hBitmap(0)
    desktopCopyBigDC.SelectBitmap desktopCopyBig

    theGrahipcs.DrawImage m_DesktopBitmapTemp.Image, 0, 0, newWidth, newHeight
    desktopCopyBitmap.hBitmap = m_DesktopBitmap.hBitmap(0)
    desktopCopySmallDC.SelectBitmap desktopCopyBitmap

    m_DesktopBitmap.CreateFromSizeFormat Me.ScaleWidth, Me.ScaleHeight, GDIPlusWrapper.Format32bppArgb
    theGrahipcs.FromImage m_DesktopBitmap.Image
    
    Dim tempImage As New GDIPBitmap
    tempImage.CreateFromHBITMAP desktopCopyBitmap.hBitmap, 0
    
    'tempImage.Image.Save "C:\cock.bmp", encoder.EncoderForMimeType("image/bmp").CodecCLSID
    
    theGrahipcs.Clear
    theGrahipcs.CompositingQuality = CompositingQualityHighQuality
    theGrahipcs.InterpolationMode = InterpolationModeHighQualityBilinear
    theGrahipcs.DrawImage tempImage.Image, 0, 0, Me.ScaleWidth + 5, Me.ScaleHeight + 5
    
    desktopCopyBitmap.hBitmap = m_DesktopBitmap.hBitmap(0) ''''''''''''' CREATING RANDOM BLACK BORDER
    desktopCopySmallDC.SelectBitmap desktopCopyBitmap
    
    BitBlt desktopCopySmallDC.Handle, 0, 0, Me.ScaleWidth, Me.ScaleHeight, m_StartMenuMaskInvertedDC.hdc, 0, 0, vbSrcAnd
    'BitBlt m_StartMenuMaskDC.Handle, 0, 0, Me.ScaleWidth, Me.ScaleHeight, m_StartMenuMaskDC.Handle, 0, 0, vbSrcInvert
    'SavePicture modTemp2.CreateBitmapPicture(desktopCopyBitmap.hBitmap), "C:\fffs.bmp"
    'm_DesktopBitmap.Image.Save "C:\fffs2.bmp", encoder.EncoderForMimeType("image/bmp").CodecCLSID
    
    theGrahipcs.Clear
    
    hDCMemory = theGrahipcs.GetHDC
    
    ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, Me.ScaleWidth, Me.ScaleHeight, desktopCopyBigDC.Handle, _
            0, 0, vbSrcCopy)
    
    ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, Me.ScaleWidth, Me.ScaleHeight, m_StartMenuMaskDC.Handle, _
            0, 0, vbSrcAnd)
    
    ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, Me.ScaleWidth, Me.ScaleHeight, desktopCopySmallDC.Handle, _
            0, 0, vbSrcPaint)
    
    theGrahipcs.ReleaseHDC hDCMemory
End Sub

Sub CloseMe()
    On Error GoTo Handler

Dim thisForm As Form

    RaiseEvent onClose
    
    m_toolTip.SetToolTip ""

    g_bStartMenuVisible = False
    m_AllPrograms.Visible = False
    
    m_programMenu.ColapseAllAndResetStrictSearch
    m_programMenu.ResetScrollbarsValues
    
    m_programMenu.Hide
    
    m_recentPrograms.ResetRollover
    m_recentPrograms.Hide
    ShowWindow m_searchText.hWnd, SW_HIDE
    
    Me.Hide
    
    m_searchText.Font = m_searchBoxForeFontItalic
    
    m_searchText.Text = GetPublicString("strStartSearch", "Start Search")
    m_AllProgramsCaption = GetInitialString
    m_ArrowAllPrograms.State = 0
    m_powerMenu.Hide
    
    m_programMenu.Filter = vbNullString
    
    timTreeViewSearch.Enabled = False
    timAutoClick.Enabled = False
    timRolloverDelay.Enabled = False
    
    If Not m_currentFadeInImage Is Nothing Then
        Unload m_currentFadeInImage
    End If
    
    Set m_currentFadeInImage = Nothing
    
    For Each thisForm In Forms
        If thisForm.Name = frmRolloverImage.Name Then
            Unload thisForm
        End If
    Next
    
    m_jumpListMode = False
    
    If m_undoLayeredOnClose Then
         m_undoLayeredOnClose = False
         
         UndoLayeredWindow
     End If
    
    ReDraw
    
    Exit Sub
Handler:
    LogError Err.Description, Me.Name
End Sub

Sub ReDraw()
    If timMorphToJumpList.Enabled Then Exit Sub
    If Not m_initalized Then Exit Sub
    
    m_BitmapGraphics.Clear
    
    If m_jumpListMode Then
    
    
        m_BitmapGraphics.DrawImage m_BackGroundPinned, 0, 0, m_BackGroundPinned.Width, m_BackGroundPinned.Height
        
        m_jumpListDrawer.Update
        m_BitmapGraphics.DrawImage m_jumpListDrawer.Image, Layout.JumpListViewerSchema.Left, Layout.JumpListViewerSchema.Top, m_jumpListDrawer.Width, m_jumpListDrawer.Height
        
    Else
        If BlurEnabled Then
            Debug.Print "Aha!!"
            m_navigationDraw.UpdateBlurred m_DesktopBitmap.Image
            m_BitmapGraphics.DrawImage m_DesktopBitmap.Image, 0, 0, Me.ScaleWidth, Me.ScaleHeight
        Else
            m_navigationDraw.UpdateLayered
        End If
        
        m_BitmapGraphics.DrawImage m_background, 0, 0, m_background.Width, m_background.Height
        m_BitmapGraphics.DrawImage m_navigationDraw.Image, Layout.GroupMenuSchema.Left, Layout.GroupMenuSchema.Top, Layout.GroupMenuSchema.Width, Layout.GroupMenuSchema.Height
    End If
    
    DrawPowerButton m_shutDownButton
    'm_bitmapgraphics.DrawString "Shutdown",
    
    DrawPowerButton m_logOffButton
    DrawPowerButton m_arrowButton
    
    
    DrawAllProgramsButton
    
    If m_ShutDownTextEnabled Then
        If Layout.ShutDownTextSchema.Visible Then
            DrawShutDownText
        End If
    End If
    
    DrawAllProgramsArrow
    DrawAllPrograms
    UpdateBuffer
End Sub

Private Sub MorphToNormal()

    ShowAvatarRollover

    m_jumpListMode = False
    m_reBlurcallback = True
    
    MorphStartMenu m_BackGroundPinned, m_background

End Sub

Private Function MorphToJumpList() As Boolean

    If Not m_currentFadeInImage Is Nothing Then
        
        If ExistInCol(m_FadeOuts, m_currentFadeInImage.Tag) = False Then
            m_FadeOuts.Add m_currentFadeInImage, m_currentFadeInImage.Tag
            Set m_currentFadeInImage = Nothing
            timFade.Enabled = True
        Else
            Exit Function
        End If
        
    End If

    MorphStartMenu m_background, m_BackGroundPinned
    If BlurEnabled Then UnblurMe False, True
    
    m_jumpListMode = True
    m_reBlurcallback = False
    
    MoveWindow Me.hWnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, m_BackGroundPinned.Width, m_BackGroundPinned.Height, False

    MorphToJumpList = True
End Function

Private Sub SetupNavigationViewer(ByRef originalBackground As GDIPImage)

    m_navigationDraw.Constructor g_resourcesPath, _
                                Layout, _
                                originalBackground, _
                                Settings.navigationXML
    
    If Layout.ForceClearType Or BlurEnabled Then
        m_navigationDraw.SetClearType
    End If

End Sub

Private Sub SetupJumplistViewer()

Dim newSize As SIZEL
Dim BackgroundPosition As POINTL

    Set m_jumpListDrawer = New JumpListDrawer
    
    newSize.cx = Layout.JumpListViewerSchema.Width
    newSize.cy = Layout.JumpListViewerSchema.Height
    
    If Layout.JumpListViewerSchema.FontID <> "" Then
        m_jumpListDrawer.Font = Layout.Fonts(Layout.JumpListViewerSchema.FontID)
    End If
    
    m_jumpListDrawer.Size = newSize
    m_jumpListDrawer.BackgroundSource = m_BackGroundPinned
    
    BackgroundPosition.X = Layout.JumpListViewerSchema.Left
    BackgroundPosition.Y = Layout.JumpListViewerSchema.Top
    
    m_jumpListDrawer.BackgroundPosition = BackgroundPosition
End Sub

Private Sub InitializePowerMenu()

    'Set callback
    Set m_powerMenu = frmVistaMenu

    m_powerMenu.AddItem UserVariable("strLogOff"), "LOGOFF"
    m_powerMenu.AddItem UserVariable("strSwitchUser"), "rundll32.exe user32.dll, LockWorkStation"
    m_powerMenu.AddItem ""
    m_powerMenu.AddItem UserVariable("strRestart"), "REBOOT"
    m_powerMenu.AddItem ""
    m_powerMenu.AddItem UserVariable("strShutdown"), "SHUTDOWN"
    m_powerMenu.AddItem ""
    m_powerMenu.AddItem UserVariable("strStandBy"), "STANDBY"
    m_powerMenu.AddItem UserVariable("strHibernate"), "HIBERNATE"
    m_powerMenu.AddItem ""
    m_powerMenu.AddItem UserVariable("strOptions"), "OPTIONS"
    m_powerMenu.AddItem UserVariable("strAbout"), "ABOUT"
    m_powerMenu.AddItem UserVariable("strExit"), "EXIT"
    
    m_powerMenu.SetDimensions
End Sub

Private Sub InitializeRecentPrograms()

Dim theFont As GDIFont
Dim defaultFont As ViFont

    If Not m_recentPrograms Is Nothing Then Unload m_recentPrograms
    
    Set m_recentPrograms = New frmFreq
    Set defaultFont = New ViFont
    
    defaultFont.Size = 9
    defaultFont.Colour = vbBlack
    defaultFont.Face = g_DefaultFont.FontFace
    
    m_recentPrograms.BackColor = CLng(Layout.FrequentProgramsMenuColour)

    If Layout.FrequentProgramsMenuSchema.FontID <> "" Then
        m_recentPrograms.TrueFont = Layout.Fonts(Layout.FrequentProgramsMenuSchema.FontID)
    Else
        m_recentPrograms.TrueFont = defaultFont
    End If
    
    SetOwner m_recentPrograms.hWnd, Me.hWnd
    
End Sub

Private Sub InitializeGroupMenu()

Dim theFont As ViFont

    m_groupMenuFontSize = 12
    m_OptionsGap = 35
    
    If Layout.GroupOptionsSeparator <> 0 Then
        m_OptionsGap = Layout.GroupOptionsSeparator
    End If

    If Layout.GroupMenuSchema.FontID <> "" Then
        Set theFont = Layout.Fonts(Layout.GroupMenuSchema.FontID)
    
        m_GroupOptionsFontFamily.Constructor theFont.Face
        m_groupMenuFontSize = theFont.Size * (1.3)
        m_GroupOptionsBrush.Colour.Value = theFont.Colour
    End If

End Sub

Private Sub InitializeSearchText()

    Set m_searchText = Nothing

Set m_searchText = New VistaSearchBox
Dim theFont As ViFont

    m_searchText.Font = g_DefaultFontItalic
    
    Set m_searchBoxForeFontItalic = g_DefaultFontItalic
    Set m_searchBoxNormalFont = g_DefaultFont
    
    m_searchText.Text = GetPublicString("strStartSearch", "Start Search")
    
    If Layout.SearchBoxSchema.BackColour <> -1 Then
        m_searchText.BackColour = Layout.SearchBoxSchema.BackColour
    End If
    
    If Layout.SearchBoxSchema.FontID <> "" Then
        Set theFont = Layout.Fonts(Layout.SearchBoxSchema.FontID)
        
        Set m_searchBoxNormalFont = theFont.ToGDIFont()
        Set m_searchBoxForeFontItalic = New GDIFont
        m_searchBoxForeFontItalic.Constructor theFont.Face, theFont.Size, APITRUE
        
        m_searchText.Font = m_searchBoxForeFontItalic
        m_searchText.ForeColour = theFont.Colour
        'MsgBox Layout.SearchBoxSchema.FontID
    End If
    
    m_searchText.ForeColour = Layout.SearchBoxForeColour
    m_searchText.FocusColour = Layout.SearchBoxFocusColour
    
    SetOwner m_searchText.hWnd, Me.hWnd

    ShowWindow m_searchText.hWnd, SW_HIDE
End Sub

Private Sub InitializeProgramMenu()
    
Dim theFont As GDIFont

    If Not m_programMenu Is Nothing Then Unload m_programMenu

    Set m_programMenu = New frmTreeView
    SetOwner m_programMenu.hWnd, Me.hWnd
    
    UpdateProgramMenu Not m_InitialStartUp
    
    m_programMenu.BackColour = CLng(Layout.ProgramMenuColour)

    If Layout.ProgramMenuSchema.FontID <> "" Then
        Set theFont = Layout.Fonts(Layout.ProgramMenuSchema.FontID).ToGDIFont()

        m_programMenu.TrueFont = theFont
        m_programMenu.ForeColour = Layout.Fonts(Layout.ProgramMenuSchema.FontID).Colour
    End If
    
    m_programMenu.SeparatorFontColour = Layout.ProgramsMenuSeperatorColour
End Sub

Private Sub InitializeShutDownText()
    'WORK

Dim theFont As ViFont

    m_ShutDownCaption = UserVariable("strShutdown")

    Set m_ShutDownBrush = New GDIPBrush
    Set m_ShutDownBrush2 = New GDIPBrush
    
    Set m_ShutDownTextFontFamily = New GDIPFontFamily
    
    m_ShutDownTextFontSize = 12
    m_ShutDownBrush2.Colour.Value = Layout.ShutDownTextJumpListColour
    
    If Layout.ShutDownButtonSchema.FontID <> "" Then
        Set theFont = Layout.Fonts(Layout.ShutDownButtonSchema.FontID)
        
        m_ShutDownTextFontFamily.Constructor theFont.Face
        m_ShutDownTextFontSize = theFont.Size * (1.3)
        m_ShutDownBrush.Colour.Value = theFont.Colour
    Else
    
        m_ShutDownTextFontFamily.Constructor OptionsHelper.PrimaryFont
        m_ShutDownBrush.Colour.Value = vbWhite
    End If

End Sub

Private Sub InitializeAllProgramsText()

Dim theFont As ViFont

    m_AllProgramsCaption = GetInitialString()

    Set m_AllProgramsBrush = New GDIPBrush
    Set m_AllProgramsFontFamily = New GDIPFontFamily
    
    m_AllProgramsFontSize = 12

    If Layout.AllProgramsTextSchema.FontID <> "" Then
        Set theFont = Layout.Fonts(Layout.AllProgramsTextSchema.FontID)
    
        m_AllProgramsFontFamily.Constructor theFont.Face
        m_AllProgramsFontSize = theFont.Size * (1.3)
        m_AllProgramsBrush.Colour.Value = theFont.Colour
    Else
        m_AllProgramsFontFamily.Constructor OptionsHelper.PrimaryFont
        m_AllProgramsBrush.Colour.Value = vbBlack
    End If
    
    m_AllPrograms.Position.X = Layout.AllProgramsRolloverSchema.Left
    m_AllPrograms.Position.Y = Layout.AllProgramsRolloverSchema.Top
End Sub

Private Sub SetPowerButtonDimensionBySchema(srcPowerButton As PowerOptionButton, srcShema As GenericViElement)

Dim imgHeight As Long
Dim imgWidth As Long

    imgHeight = srcPowerButton.Image.Height / 3
    imgWidth = srcPowerButton.Image.Width
    
    With srcPowerButton.NormalButton
        .Width = imgWidth
        .Height = imgHeight
    End With
    
    With srcPowerButton.RolloverButton
        .Top = imgHeight
        .Width = imgWidth
        .Height = imgHeight
    End With
    
    With srcPowerButton.PressedButton
        .Top = imgHeight * 2
        .Width = imgWidth
        .Height = imgHeight
    End With
    
    srcPowerButton.Position.X = srcShema.Left
    srcPowerButton.Position.Y = srcShema.Top

End Sub

Private Sub SetPowerButtonDimensions()

    SetPowerButtonDimensionBySchema m_shutDownButton, Layout.ShutDownButtonSchema
    SetPowerButtonDimensionBySchema m_logOffButton, Layout.LogOffButtonSchema
    SetPowerButtonDimensionBySchema m_arrowButton, Layout.ArrowButtonSchema
    
    m_arrowButton.ActionCode = 1
    m_shutDownButton.ActionCode = 2
    m_logOffButton.ActionCode = 3

End Sub

Sub UpdateBuffer()
On Error GoTo Handler

    If Not m_layeredMode Then
    
        m_graphics.DrawImage _
            m_Bitmap.Image, 0, 0, Me.ScaleWidth, Me.ScaleHeight
            
        RepaintWindow Me.hWnd
    Else
    
        m_graphics.Clear
        
        m_graphics.DrawImage _
            m_Bitmap.Image, 0, 0, Me.ScaleWidth, Me.ScaleHeight
            
        Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, m_layeredData.GetSize, m_layeredData.theDC, m_layeredData.GetPoint, 0, m_blendFunc32bpp, ULW_ALPHA)
    End If
        
    Exit Sub
Handler:
    Debug.Print "UpdateBuffer()" & Err.Description
End Sub

Function ReInitSurface() As Boolean
    On Error GoTo Handler
    
    'm_winSize.cx = Me.ScaleWidth
    'm_winSize.cy = Me.ScaleHeight
    
    m_Bitmap.Dispose
    m_BitmapGraphics.Dispose

    m_Bitmap.CreateFromSizeFormat Me.ScaleWidth, Me.ScaleHeight, GDIPlusWrapper.Format32bppArgb
    m_BitmapGraphics.FromImage m_Bitmap.Image
    
    m_BitmapGraphics.TextRenderingHint = TextRenderingHintSingleBitPerPixelGridFit
    m_BitmapGraphics.SmoothingMode = SmoothingModeHighQuality
    m_BitmapGraphics.InterpolationMode = InterpolationModeHighQualityBicubic
    m_BitmapGraphics.PixelOffsetMode = PixelOffsetModeHighQuality

    If Not m_layeredData Is Nothing Then
        Set m_layeredData = Nothing
        Set m_layeredData = MakeLayerdWindow(Me)

        m_graphics.FromHDC m_layeredData.theDC
    Else
        m_graphics.FromHDC Me.hdc
    End If
    
    ReInitSurface = True
    
    Exit Function
Handler:
    ReInitSurface = False
    Debug.Print "ReInitSurface():" & Me.ScaleWidth & " / " & Me.ScaleHeight & vbCrLf & Err.Description
End Function

Private Sub DrawAllProgramsArrow()
    If m_ArrowAllPrograms.Visible = False Then
        Exit Sub
    End If

Dim rArrow As gdiplus.RECTL
Dim sourceRect As gdiplus.RECTL

    rArrow.Left = m_ArrowAllPrograms.Position.X
    rArrow.Top = m_ArrowAllPrograms.Position.Y
    
    If m_ArrowAllPrograms.State = 0 Then
        CopyRectL m_ArrowAllPrograms.NormalButton, sourceRect
    ElseIf m_ArrowAllPrograms.State = 1 Then
        CopyRectL m_ArrowAllPrograms.RolloverButton, sourceRect
    End If
    
    rArrow.Width = sourceRect.Width
    rArrow.Height = sourceRect.Height

    m_BitmapGraphics.DrawImageStretchAttrL m_ArrowAllPrograms.Image, _
        rArrow, _
        sourceRect.Left, sourceRect.Top, sourceRect.Width, sourceRect.Height, UnitPixel, 0, 0, 0
        
End Sub

Private Sub DrawAllProgramsButton()
    If m_AllPrograms.Visible = False Then
        Exit Sub
    End If
    
Dim r As gdiplus.RECTL

    
    r.Left = m_AllPrograms.Position.X
    r.Top = m_AllPrograms.Position.Y
    
    r.Width = m_AllPrograms.Image.Width
    r.Height = m_AllPrograms.Image.Height
    
    m_BitmapGraphics.DrawImageStretchAttrL m_AllPrograms.Image, _
        r, _
        0, 0, m_AllPrograms.Image.Width, m_AllPrograms.Image.Height, UnitPixel, 0, 0, 0
        

End Sub

Private Sub DrawPowerButton(ByRef sourceButton As PowerOptionButton)
Dim r As gdiplus.RECTL
Dim sourceRect As gdiplus.RECTL

    If Not sourceButton.Visible Then Exit Sub
    
    r.Left = sourceButton.Position.X
    r.Top = sourceButton.Position.Y
    
    If sourceButton.State = 0 Then
        CopyRectL sourceButton.NormalButton, sourceRect
    ElseIf sourceButton.State = 1 Then
        CopyRectL sourceButton.RolloverButton, sourceRect
    ElseIf sourceButton.State = 2 Then
        CopyRectL sourceButton.PressedButton, sourceRect
    End If
    
    r.Width = sourceRect.Width
    r.Height = sourceRect.Height
    
    m_BitmapGraphics.DrawImageStretchAttrL sourceButton.Image, _
        r, _
        sourceRect.Left, sourceRect.Top, sourceRect.Width, sourceRect.Height, UnitPixel, 0, 0, 0


End Sub

Private Sub DrawShutDownText()
    If m_ShutDownTextFontFamily Is Nothing Then Exit Sub

    m_Path.Constructor FillModeWinding
    m_Path.AddString m_ShutDownCaption, m_ShutDownTextFontFamily, FontStyleRegular, m_ShutDownTextFontSize, CreateRectF(CSng(Layout.ShutDownTextSchema.Left), CSng(Layout.ShutDownTextSchema.Top), 20, 130), 0
    
    If m_jumpListMode Then
        m_BitmapGraphics.FillPath m_ShutDownBrush2, m_Path
    Else
        m_BitmapGraphics.FillPath m_ShutDownBrush, m_Path
    End If

End Sub

Private Sub DrawAllPrograms()
    If m_AllProgramsFontFamily Is Nothing Then Exit Sub

    m_Path.Constructor FillModeWinding
    m_Path.AddString m_AllProgramsCaption, m_AllProgramsFontFamily, FontStyleRegular, m_AllProgramsFontSize, CreateRectF(CSng(Layout.AllProgramsTextSchema.Left), CSng(Layout.AllProgramsTextSchema.Top), 20, 130), 0
    m_BitmapGraphics.FillPath m_AllProgramsBrush, m_Path

End Sub

Private Sub HideLeftRollover()

    If g_KeyboardMenuState <> 0 Then
        g_KeyboardMenuState = 0
        m_AllPrograms.Visible = False
        
        ReDraw
    End If

End Sub

Private Sub Form_DragDropFolder(szFolderPath As String)
    If m_navigationDraw Is Nothing Then Exit Sub
    
    m_navigationDraw.AddNewFolder szFolderPath
End Sub

Private Sub Form_Initialize()
    Inititalize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Debug.Print "frmStartMenuBase:: " & KeyCode & " <> " & g_KeyboardMenuState & " <> " & g_KeyboardSide

    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
    
        If g_KeyboardMenuState = 1 Then
            If g_KeyboardSide = 1 Then
                HideLeftRollover
                ShowRightRollover m_navigationDraw.count
                
                g_KeyboardMenuState = 1
                g_KeyboardSide = 2
            ElseIf g_KeyboardSide = 2 Then

                If m_navigationDraw.index < m_navigationDraw.count Then
                    If m_recentPrograms.Visible Then
                    
                        Debug.Print "Attacking Recent Programs!"
                        m_recentPrograms.RolloverWithKeyboard (m_navigationDraw.index - 1)
                    Else
                        m_programMenu.SelectVisibleItem m_navigationDraw.index
                    End If
                    
                    g_KeyboardMenuState = 0
                    g_KeyboardSide = 1
                Else
                    g_KeyboardSide = 1
                    g_KeyboardMenuState = 1
                    m_AllPrograms.Visible = True
                    
                    ReDraw
                    
                End If
                
                m_navigationDraw.ResetRollover
                'HideLeftRollover
                'ShowRightRollover m_optionsMax
                

            End If
        Else
            SendKeyToActiveWindow CLng(KeyCode)
        End If
    
        KeyCode = 0
        Exit Sub
    End If

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        
        If g_KeyboardSide = 1 Then
        
            'SendKeyToActiveWindow CLng(KeyCode)
            If g_KeyboardMenuState = 1 Then
                g_KeyboardMenuState = 0
                m_AllPrograms.Visible = False
                
                ReDraw
                
                If m_recentPrograms.Visible Then
                    If KeyCode = vbKeyUp Then
                        m_recentPrograms.SelectBottomMost
                    ElseIf KeyCode = vbKeyDown Then
                        m_recentPrograms.SelectTopMost
                    End If
                    
                    'win.SetFocus m_recentPrograms.hWnd
                Else
                    If KeyCode = vbKeyUp Then
                        m_programMenu.SelectLastItem
                    ElseIf KeyCode = vbKeyDown Then
                        m_programMenu.SelectFirstVisibleItem
                    End If
                End If
            Else
                SendKeyToActiveWindow CLng(KeyCode)
            End If
            
        ElseIf g_KeyboardSide = 2 Then
        
            If KeyCode = vbKeyUp Then
                m_navigationDraw.SelectOptionAbove
            ElseIf KeyCode = vbKeyDown Then
                m_navigationDraw.SelectOptionBelow
            End If
        
        End If
    Else
        
        'm_searchText.SetKeyboardFocus
        'SetKeyDown CLng(KeyCode)
        
    End If

End Sub

Private Sub SendKeyToActiveWindow(KeyCode As Long)

Dim activeWindow As Long

    If m_recentPrograms.Visible Then
        activeWindow = m_recentPrograms.hWnd
    Else
        activeWindow = m_programMenu.hWnd
    End If
    
    PostMessage activeWindow, WM_KEYDOWN, ByVal KeyCode, 0
    PostMessage activeWindow, WM_KEYUP, ByVal KeyCode, 0

    'win.SetFocus activeWindow
End Sub

Private Function GetInitialString() As String

    If Not Settings.ShowProgramsFirst Then
        GetInitialString = GetPublicString("strAllPrograms", "All Programs")
    Else
        GetInitialString = "Frequent Programs"
    End If

End Function

Private Sub ToggleProgramsMenu()
    
    If m_AllProgramsCaption = GetInitialString Then
        m_AllProgramsCaption = GetPublicString("strBack", "Back")
        m_ArrowAllPrograms.State = 1
    Else
        m_AllProgramsCaption = GetInitialString
        m_ArrowAllPrograms.State = 0
    End If

    If m_recentPrograms.Visible Then
    
        If m_jumpListMode Then
            m_recentPrograms.ResetRollover
            MorphToNormal
        End If
        
        UpdateProgramMenuIfRequired
        
        m_programMenu.ColapseAllAndResetStrictSearch
        m_programMenu.ResetScrollbarsValues
        m_programMenu.ResetKeyboardStatus
        
        m_recentPrograms.Hide
    Else
        m_recentPrograms.Show
        
    End If
    
    win.SetFocus Me.hWnd
    ReDraw
End Sub

Private Sub UpdateProgramMenu(Optional refreshIconCache As Boolean = False)

    ProgramIndexingHelper.Initialize
    
    CleanCollection m_programMenu.RootNode.Children
    CleanCollection m_programMenu.RootPrograms
    
    'PopulateNode m_programMenu.RootNode, FSO.GetFolder("E:\StarTrek TNG")
    PopulateStarMenuNodes m_programMenu.RootNode
    
    m_programMenu.ForceRepaint

End Sub

Private Sub UpdateProgramMenuIfRequired()

    If FileCountTest = True Then
        UpdateProgramMenu
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Debug.Print "KeyCode:: " & KeyCode

    If g_KeyboardMenuState = 1 Then
        
        If KeyCode = vbKeyReturn Then
    
            If g_KeyboardSide = 1 Then
                If m_searchText.Text <> vbNullString Then
                    ShellEx m_searchText.Text
                Else
                    ToggleProgramsMenu
                End If
            ElseIf g_KeyboardSide = 2 Then
                'ExecuteRolloverCommand m_options(m_currentRolloverIndex - 1).Command
            End If
        End If
    End If

End Sub

Sub Inititalize()

    m_recentItemsAutoClick = True
    m_InitialStartUp = True
    
   m_AlphaAmount = 1
            
    With m_blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    Set g_AutomaticDestinationsUpdater = New AutomaticDestinationsUpdater
    timJumplistUpdater.Enabled = g_AutomaticDestinationsUpdater.JumplistsAvailable
    
    Set m_FadeOuts = New Collection
    Set m_optionDialog = frmControlPanel
    Set m_jumpListDrawer = New JumpListDrawer
    Set m_navigationDraw = Settings.NavigationPane
    
    Set m_graphics = New GDIPGraphics
    Set m_BitmapGraphics = New GDIPGraphics
    Set m_BlackBrush = New GDIPBrush
    
    Set m_DesktopBitmap = New GDIPBitmap
    Set m_StartMenuMaskBitmap = New GDIBitmap
    Set m_StartMenuMaskDC = New GDIDC
    Set m_StartMenuMaskInvertedDC = New pcMemDC
    Set m_Bitmap = New GDIPBitmap
    
    Set m_RecentItems = New frmFileMenu
    m_navigationDraw.BaseWindow = Me
    
    m_BlackBrush.Colour.SetColourByHex "#000000"
    
    InitializePowerMenu
    InitializeCurrentSkin
    
    Call HookWindow(Me.hWnd, Me)
    KeyBoard.HookKeyboard Me.hWnd
    
    m_InitialStartUp = False

    DragAcceptFiles Me.hWnd, APITRUE
    
    AddToShellContextMenu "lnkfile"
    AddToShellContextMenu "*"
        
    If m_toolTip Is Nothing Then Set m_toolTip = New ViToolTip
    m_toolTip.AttachWindow Me.hWnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_showingMenu Is Nothing Then
        If m_showingMenu.Visible Then
            Exit Sub
        End If
    End If
    
    timAutoClick.Enabled = False
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, XSng As Single, YSng As Single)
    If m_recentPrograms Is Nothing Then Exit Sub
    If g_Exiting Then Exit Sub

Static lastPosition As POINTS
Static lastButton As Integer

Dim clientMousePos As POINTL
Dim clicked As Boolean

    If lastPosition.X = XSng And lastPosition.Y = YSng And lastButton = Button Then
        Exit Sub
    End If
    
    If lastButton = 0 And Button <> 0 Then
        clicked = True
    End If
    
    lastButton = Button
    
    lastPosition.X = XSng
    lastPosition.Y = YSng
    
    If Not m_showingMenu Is Nothing Then
        If m_showingMenu.Visible Then
            Exit Sub
        End If
    End If

    If m_jumpListMode Then
        If IsInsideViComponent(XSng, YSng, Layout.JumpListViewerSchema, clientMousePos) Then
            m_jumpListDrawer.HasMouse = True
            
            If Button = vbLeftButton Then
                If clicked Then m_jumpListDrawer.MouseLeftClick CLng(XSng), CLng(YSng)
            ElseIf Button = vbRightButton Then
                If clicked Then m_jumpListDrawer.MouseRightClick CLng(XSng), CLng(YSng)
            Else
                m_jumpListDrawer.MouseMove clientMousePos
            End If
            
            Exit Sub
        ElseIf m_jumpListDrawer.HasMouse Then
        
            m_jumpListDrawer.HasMouse = False
            m_jumpListDrawer.MouseLeaves
        End If
    Else
        Debug.Print XSng & ":" & Layout.GroupMenuSchema.Left & " & " & Layout.GroupMenuSchema.Width
    
        If IsInsideViComponent(XSng, YSng, Layout.GroupMenuSchema, clientMousePos) Then
            m_navigationDraw.HasMouse = True
        
            If Button = vbLeftButton Then
                m_navigationDraw.MouseLeftClick
            ElseIf m_navigationDraw.LeftMouseButtonDown And Button = 0 Then
                m_navigationDraw.MouseUp
            Else
                m_navigationDraw.MouseMove clientMousePos
            End If
        ElseIf m_navigationDraw.HasMouse Then
        
            m_navigationDraw.HasMouse = False
            m_navigationDraw.MouseOut
        Else
        
            If Button = vbRightButton Then
                
            
                If Not m_showingMenu Is Nothing Then
                    m_showingMenu.Hide
                End If
            
                If Not m_contextMenu Is Nothing Then
                    Set m_showingMenu = m_contextMenu
                    m_contextMenu.Resurrect True, Me
                End If
            End If
        
        End If
    End If
    
    If m_trackingMouse = False Then
        m_trackingMouse = TrackMouse(Me.hWnd)
        
        m_recentPrograms.TestRolloverVisability
    End If
    
    If m_ignoreActivation And Button = vbLeftButton Then
        m_ignoreActivation = False
        Exit Sub
    End If

Dim hasChanged   As Boolean
Dim X As Long:              X = CLng(XSng)
Dim Y As Long:              Y = CLng(YSng)
Dim currentPoint As POINTL: currentPoint = CreatePointL(Y, X)
    
    hasChanged = False
    
    'If Not m_jumpListMode Then UpdateRolloverState Button, X, Y, hasChanged
    UpdatePowerButtonStatus Button, X, Y, hasChanged
    UpdateAllProgramsRollover Button, X, Y, hasChanged

    timAutoClick.Enabled = False

    'If m_recentItemsAutoClick Then
        'UpdateAutoClickButton currentPoint, CreateRect(Layout.GroupMenuSchema.Left, _
                                                                         m_recentItems_Y, _
                                                                         Layout.GroupMenuSchema.Left + m_rollover.NormalButton.Width, _
                                                                         m_recentItems_Y + m_rollover.NormalButton.Height)
        
    'End If

    If m_AllPrograms.AutoClick Then
        UpdateAutoClickButton currentPoint, CreateRect(m_AllPrograms.Position.X, _
                                                                         m_AllPrograms.Position.Y, _
                                                                         m_AllPrograms.Position.X + m_AllPrograms.Image.Width, _
                                                                         m_AllPrograms.Position.Y + m_AllPrograms.Image.Height)
    End If
    
    If m_arrowButton.AutoClick Then
        UpdateAutoClickButton currentPoint, CreateRect(m_arrowButton.Position.X, _
                                                                         m_arrowButton.Position.Y, _
                                                                         m_arrowButton.Position.X + m_arrowButton.NormalButton.Width, _
                                                                         m_arrowButton.Position.Y + m_arrowButton.NormalButton.Height)
    End If
    
    If hasChanged Then
        ReDraw
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim clientMousePos As POINTL
    
    If Not m_showingMenu Is Nothing Then
        If m_showingMenu.Visible Then
            Exit Sub
        End If
    End If
    
    ExecuteCurrentAllPrograms CLng(Button), X, Y
    
    If Button = vbRightButton Then
        If IsInsideViComponent(X, Y, Layout.GroupMenuSchema, clientMousePos) Then
            m_navigationDraw.ShowContextMenu Me
        End If
    End If
    
    'Resets all buttons to normal (non-clicked states)
    Form_MouseMove 0, Shift, X, Y
End Sub

Private Sub Form_Resize()
    If Not m_initalized Then Exit Sub
    
    ReInitSurface
    ReDraw
End Sub

Private Sub ExecuteCurrentAllPrograms(mouseButton As MouseButtonConstants, X As Single, Y As Single)

    If m_AllPrograms.State = 1 And mouseButton = vbLeftButton Then
        UpdateProgramMenuIfRequired
    
        ToggleProgramsMenu
        ReDraw
    End If

End Sub

Private Sub PowerButtonAction(sourceButtion As PowerOptionButton)
    
Dim ActionCode As Long
    
    ActionCode = sourceButtion.ActionCode

    If ActionCode = 1 Then
        
        sourceButtion.AllowUpdates = False
    
        If m_powerMenu Is Nothing Then
            'Set callback
            InitializePowerMenu
        End If

        'frmVistaMenu fits on the screen
        If (Me.Left + Me.Width + m_powerMenu.Width) > Screen.Width Then
        
            'frmVistaMenu.Move (sourceButtion.Position.x * Screen.TwipsPerPixelX) - frmVistaMenu.Width, _

            MoveWindow m_powerMenu.hWnd, _
                        (m_windowPosition.X + sourceButtion.Position.X) - m_powerMenu.ScaleWidth, _
                        (m_windowPosition.Y + sourceButtion.Position.Y) - m_powerMenu.ScaleHeight, _
                        m_powerMenu.ScaleWidth, m_powerMenu.ScaleHeight, True

        Else
            'frmVistaMenu.Move (sourceButtion.Position.x * Screen.TwipsPerPixelX), _
                        ((sourceButtion.Position.y * Screen.TwipsPerPixelY))
                        
                        
            MoveWindow m_powerMenu.hWnd, _
                        m_windowPosition.X + sourceButtion.Position.X, _
                        (m_windowPosition.Y + sourceButtion.Position.Y) - POWERMENU_HEIGHT, _
                        POWERMENU_WIDTH, POWERMENU_HEIGHT, True
                            
        End If
        
        m_powerMenu.Tag = 0
        m_powerMenu.Resurrect False, Me
        Set m_showingMenu = m_powerMenu
        
    ElseIf ActionCode = 2 Then
        PowerHelper.ExitWindowsEx EWX_POWEROFF, EWX_FORCEIFHUNG
    ElseIf ActionCode = 3 Then
        PowerHelper.ExitWindowsEx EWX_LOGOFF, EWX_FORCEIFHUNG
    End If

End Sub

Private Sub UpdatePowerButtonStatus(mouseButton As Integer, X As Long, Y As Long, ByRef hasChanged As Boolean)

    UpdatePowerButtonState m_shutDownButton, hasChanged, mouseButton, X, Y
    UpdatePowerButtonState m_logOffButton, hasChanged, mouseButton, X, Y
    UpdatePowerButtonState m_arrowButton, hasChanged, mouseButton, X, Y
End Sub

Private Function UpdatePowerButtonState(ByRef sourceButton As PowerOptionButton, ByRef hasChanged As Boolean, mouseButton As Integer, ByRef X As Long, ByRef Y As Long) As Boolean
    If Not sourceButton.AllowUpdates Or Not sourceButton.Visible Then
        Exit Function
    End If

    If X > sourceButton.Position.X And X < (sourceButton.Position.X + sourceButton.NormalButton.Width) And _
        Y > sourceButton.Position.Y And Y < (sourceButton.Position.Y + sourceButton.NormalButton.Height) Then

        If mouseButton = vbLeftButton Then
            If sourceButton.State <> 2 Then
                hasChanged = True
                sourceButton.State = 2
                
                sourceButton.AutoClick = False
                PowerButtonAction sourceButton
            End If
            
        ElseIf mouseButton = 0 Then
        
            If sourceButton.State <> 1 Then
                hasChanged = True
                sourceButton.State = 1
            End If
        End If
    Else
        sourceButton.AutoClick = True
    
        If sourceButton.State <> 0 Then
            hasChanged = True
            sourceButton.State = 0
            
            
        End If
    End If
End Function

Private Sub ShowRightRollover(index As Long)
    Debug.Print "ShowRightRollover; " & index
    
    m_navigationDraw.SelectOptionByIndex index
End Sub

Private Sub ShowAvatarRollover(Optional DefaultAlpha As Byte = 45)
    'MsgBox "ShowAvatarRollover:: " & DefaultAlpha

    If g_bStartMenuVisible = False Or _
        g_rolloverImage Is Nothing Then
        Exit Sub
    End If
    
    'MsgBox "S " & DefaultAlpha
    
    If Not m_currentFadeInImage Is Nothing Then
        If m_currentFadeInImage.Path = "SYS_ROLLOVER" Then
            Exit Sub
        End If
        Set m_currentFadeOutImage = m_currentFadeInImage
        
        If ExistInCol(m_FadeOuts, m_currentFadeOutImage.Tag) = False Then
            m_FadeOuts.Add m_currentFadeOutImage, m_currentFadeOutImage.Tag
        End If
    End If
    
    'MsgBox "4 " & DefaultAlpha
    Set m_currentFadeInImage = New frmRolloverImage
    Set m_userPicture = m_currentFadeInImage
    
    If Settings.ShowUserPicture = False Then Exit Sub
    
    m_currentFadeInImage.MakeTransFromImage MainHelper.g_rolloverImage
    m_currentFadeInImage.Path = "SYS_ROLLOVER"

    SetWindowPos m_currentFadeInImage.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
    SetOwner m_currentFadeInImage.hWnd, Me.hWnd
 
    m_currentFadeInImage.Alpha = DefaultAlpha
    
    MoveWindow m_currentFadeInImage.hWnd, m_userPicturePlaceHolder.X, m_userPicturePlaceHolder.Y, m_currentFadeInImage.ScaleWidth, m_currentFadeInImage.ScaleHeight, False
    ShowWindow m_currentFadeInImage.hWnd, SW_SHOWNA
    
    timFade.Enabled = True
End Sub

Private Sub ShowRollover(ByVal Path As String, Optional DefaultAlpha As Byte = 45)
    If g_bStartMenuVisible = False Then
        Exit Sub
    End If
    
    If Not m_currentFadeInImage Is Nothing Then
        If m_currentFadeInImage.Path = Path Then
            Exit Sub
        End If
    
        'If Not m_currentFadeOutImage Is Nothing Then
            'If ExistInCol(m_FadeOuts, m_currentFadeOutImage.Tag) = False Then
                'm_FadeOuts.Add m_currentFadeOutImage, m_currentFadeOutImage.Tag
                'Debug.Print "Adding bad fadeout: " & m_currentFadeOutImage.Tag & "<>" & m_currentFadeOutImage.Path
            'End If
        'End If
        
        If Not m_currentFadeOutImage Is Nothing Then
            Unload m_currentFadeOutImage
        End If
        
        Set m_currentFadeOutImage = m_currentFadeInImage
        
        If ExistInCol(m_FadeOuts, m_currentFadeOutImage.Tag) = False Then
            m_FadeOuts.Add m_currentFadeOutImage, m_currentFadeOutImage.Tag
            Debug.Print "Add fadeout: " & m_currentFadeOutImage.Tag
        End If
        
        'unload m_currentFadeInImage
    End If
    
    Set m_currentFadeInImage = GenerateRolloverImage(Path)
    'm_currentFadeInImage.Hide
    
    SetWindowPos m_currentFadeInImage.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
    SetOwner m_currentFadeInImage.hWnd, Me.hWnd
    
    m_currentFadeInImage.Alpha = DefaultAlpha
    
    MoveWindow m_currentFadeInImage.hWnd, m_rolloverImagesPlaceHolder.X, m_rolloverImagesPlaceHolder.Y, m_currentFadeInImage.ScaleWidth, m_currentFadeInImage.ScaleHeight, False
    ShowWindow m_currentFadeInImage.hWnd, SW_SHOWNA
    
    timFade.Enabled = True
End Sub

Private Sub UpdateAllProgramsRollover(mouseButton As Integer, X As Long, Y As Long, hasChanged As Boolean)

Dim Visible As Boolean
Dim State As Long

    If X > m_AllPrograms.Position.X And X < m_AllPrograms.Position.X + m_AllPrograms.Image.Width And _
        Y > m_AllPrograms.Position.Y And Y < m_AllPrograms.Position.Y + m_AllPrograms.Image.Height Then
        
        Visible = True
        
        If mouseButton = vbLeftButton Then
            State = 1
            m_AllPrograms.AutoClick = False
        Else
            State = 0
        End If
                
    End If
    
    If Not Visible Then
        'g_KeyboardMenuState = 0
        m_AllPrograms.AutoClick = True
    Else
        g_KeyboardMenuState = 1
        g_KeyboardSide = 1
        
        m_programMenu.ResetKeyboardStatus
        
    End If
    
    If m_AllPrograms.Visible <> Visible Then
        m_AllPrograms.Visible = Visible
        hasChanged = True
    End If

    If m_AllPrograms.State <> State Then
        Debug.Print "STATE CHANGED; " & State

        m_AllPrograms.State = State
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindow Me.hWnd
    UnhookKeyboard
    
    If Not m_recentPrograms Is Nothing Then Unload m_recentPrograms
    If Not m_programMenu Is Nothing Then Unload m_programMenu
    If Not m_powerMenu Is Nothing Then Unload m_powerMenu
    If Not m_RecentItems Is Nothing Then Unload m_RecentItems
    
    Set m_searchText = Nothing
    
    RemoveFromShellContextMenu "*"
    RemoveFromShellContextMenu "lnkfile"
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long

    On Error GoTo Handler

    If msg = WM_ACTIVATEAPP Then
        If LOWORD(wp) = WA_INACTIVE Then
            If g_bStartMenuVisible And _
                (Not MouseInsideWindow(g_ViGlanceOrb)) Then
                
                Debug.Print "CloseMe::WA_INACTIVE:: " & lp & " " & App.ThreadID
                CloseMe
            End If
        End If
    ElseIf msg = WM_MOVE Then
        'm_programMenu.Left = Me.Left + (10 * Screen.TwipsPerPixelX)
        'm_programMenu.Top = Me.Top
        
        m_windowPosition.X = LOWORD(lp)
        m_windowPosition.Y = HiWord(lp)
        
        'AlignChildWindows
    
        
    ElseIf msg = WM_ACTIVATE Then
        
        If wp = WA_ACTIVE Then
            'm_searchText.SetKeyboardFocus
        Else
            ' Just allow default processing for everything else.
            IHookSink_WindowProc = _
                CallOldWindowProcessor(hWnd, msg, wp, lp)
        End If
    
    ElseIf msg = WM_MOUSEWHEEL Then
        
        PostMessage m_programMenu.hWnd, WM_MOUSEWHEEL, ByVal wp, 0

    ElseIf msg = WM_MOUSELEAVE Then
    
        Form_MouseMove 0, 0, 0, 0
        m_trackingMouse = False
        timAutoClick.Enabled = False
        timRolloverDelay.Enabled = False
        
        
        If Not m_jumpListMode Then ShowAvatarRollover

    ElseIf msg = UM_CLOSE_STARTMENU Then
    
        Debug.Print "CloseMe::UM_CLOSE_STARTMENU"
        CloseMe
    
    ElseIf msg = WM_ENDSESSION Then
    
        ExitApplication
        IHookSink_WindowProc = APITRUE
        
    ElseIf msg = WM_DROPFILES Then

        Dim hFilesInfo As Long
        Dim szFileName As String
        Dim wTotalFiles As Long
        Dim wIndex As Long
        
        hFilesInfo = wp
        wTotalFiles = DragQueryFileW(hFilesInfo, &HFFFF, ByVal 0&, 0)
    
        For wIndex = 0 To wTotalFiles
            szFileName = Space$(1024)
            
            If Not DragQueryFileW(hFilesInfo, wIndex, StrPtr(szFileName), Len(szFileName)) = 0 Then
                Form_DragDropFolder TrimNull(szFileName)
            End If
        Next wIndex
        
        DragFinish hFilesInfo
        
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

Private Function AlignChildWindows() As Boolean
    On Error GoTo Handler

    MoveWindow m_programMenu.hWnd, m_windowPosition.X + Layout.ProgramMenuSchema.Left, _
                                  m_windowPosition.Y + Layout.ProgramMenuSchema.Top, _
                                  Layout.ProgramMenuSchema.Width, _
                                  Layout.ProgramMenuSchema.Height, True
                                  
    MoveWindow m_recentPrograms.hWnd, m_windowPosition.X + Layout.FrequentProgramsMenuSchema.Left, _
                                     m_windowPosition.Y + Layout.FrequentProgramsMenuSchema.Top, _
                                    Layout.FrequentProgramsMenuSchema.Width, _
                                    Layout.FrequentProgramsMenuSchema.Height, True
    
    MoveWindow m_searchText.hWnd, m_windowPosition.X + Layout.SearchBoxSchema.Left, _
                                  m_windowPosition.Y + Layout.SearchBoxSchema.Top, _
                                  Layout.SearchBoxSchema.Width, _
                                  Layout.SearchBoxSchema.Height, True

    m_rolloverImagesPlaceHolder.X = m_windowPosition.X + Layout.RolloverPlaceHolder.Left
    m_rolloverImagesPlaceHolder.Y = m_windowPosition.Y + Layout.RolloverPlaceHolder.Top

    m_userPicturePlaceHolder.X = m_windowPosition.X + Layout.UserPictureSchema.Left
    m_userPicturePlaceHolder.Y = m_windowPosition.Y + Layout.UserPictureSchema.Top
    
    AlignChildWindows = True
    Exit Function
Handler:
    LogError Err.Description, "AlignChildWindows()"
    AlignChildWindows = False
End Function

Private Sub m_jumpListDrawer_onChanged(newItem As JumpListItem)
    ReDraw
    
    Debug.Print "Setting tooltip:: " & newItem.Caption
    m_toolTip.SetToolTip newItem.Path
    m_toolTip.Hide
End Sub

Private Sub m_jumpListDrawer_onMouseExit()
    m_toolTip.SetToolTip ""
    m_toolTip.Hide
    
    ReDraw
End Sub

Private Sub m_jumpListDrawer_onRequestClose()
    CloseMe
End Sub

Private Sub m_navigationDraw_onChanged()
    ReDraw
    
    If Not m_navigationDraw.SelectedItem Is Nothing Then
        m_RolloverToShow = g_rolloverPath & m_navigationDraw.SelectedItem.Rollover
        If FileExists(m_RolloverToShow) Then
            timRolloverDelay.Enabled = False
            timRolloverDelay.Enabled = True
        Else
            ShowAvatarRollover
        End If
    End If
End Sub

Private Sub m_navigationDraw_onCommand(szCommand As String)

    If szCommand <> "" Then
        If Left$(szCommand, 1) = ":" Then
            'Use this to position spawned Run window
            'g_WinLib.Move g_winStartMenuBase.Left + 500, g_winStartMenuBase.Top + (g_winStartMenuBase.Height / 2)

            'MsgBox "DEBUG::STARTMENU"
            Dim dummyWin As New frmZOrderKeeper
            MoveWindow dummyWin.hWnd, m_windowPosition.X + 33, m_windowPosition.Y + (Me.ScaleHeight / 2), 1, 1, False
            
            If LCase$(Right$(szCommand, Len(szCommand) - 1)) = "filerun" Then
                CloseMe
                ShowRun dummyWin.hWnd
            ElseIf LCase$(Right$(szCommand, Len(szCommand) - 1)) = "findfiles" Then
                CloseMe
                ShowFind
            Else
                'ShowRecentItems
            End If
            
            Unload dummyWin
            Set dummyWin = Nothing
        Else
            CloseMe
            
            If InStr(LCase$(szCommand), "helppane.exe") > 0 Then
                AppLauncherHelper.CmdRun szCommand, vbHide
                AppLauncherHelper.CmdRun szCommand, vbHide
            Else
                AppLauncherHelper.CmdRun szCommand
            End If
        End If
    End If


End Sub

Private Sub m_optionDialog_onRequestRecentProgramsRefresh()
    m_recentPrograms.PopulateItems
End Sub

Private Sub m_powerMenu_onCommand(commandCode As PowerMenuCommands)
    
    CloseMe
    
    If commandCode = ShowOptions Then
        m_optionDialog.Show
        
    ElseIf commandCode = ShowAbout Then
        m_optionDialog.Show
        m_optionDialog.NavigateToPanel "about"
        
    ElseIf commandCode = Reboot Then
        frmZOrderKeeper.InitiateRestart
        
    ElseIf commandCode = PowerOff Then
        frmZOrderKeeper.InitiateShutDown
        
    ElseIf commandCode = LogOff Then
        frmZOrderKeeper.InitiateLogOff
        
    ElseIf commandCode = Hibernate Then
        SetSuspendState True, True, False
    ElseIf commandCode = StandBy Then
        SetSuspendState False, True, False
    End If
    
End Sub

Private Sub m_optionDialog_onChangeSkin(szNewSkin As String)
    Me.Skin = szNewSkin
End Sub

Private Sub m_optionDialog_onNavigationPanelChange()
    m_navigationDraw.NotifyOptionsChanged
    ReDraw
End Sub

Private Sub m_optionDialog_onRequestAddMetroShortcut()
    Settings.Programs.AddMetroShortcut_ToPinned
    If g_Windows8 or g_Windows81 Then Settings.Programs.AddMetroAppsShortcut_ToPinned
End Sub

Private Sub m_powerMenu_onClick(theItemTag As String)

Dim sP() As String
Dim lngIndex As Long

    Select Case theItemTag
    
        Case "SEP"
            'Do nothing
        
        Case "REBOOT"
            'frmStartMenuBase.CloseMe
            'PowerMod.ExitWindowsEx EWX_REBOOT, EWX_FORCEIFHUNG
            
            PutOptions
            m_powerMenu_onCommand Reboot
        
        Case "SHUTDOWN"
            'frmStartMenuBase.CloseMe
            'PowerMod.ExitWindowsEx EWX_POWEROFF, EWX_FORCEIFHUNG
            
            PutOptions
            m_powerMenu_onCommand PowerOff
            
        Case "LOGOFF"
            'frmStartMenuBase.CloseMe
            'PowerMod.ExitWindowsEx EWX_LOGOFF, EWX_FORCEIFHUNG
            
            PutOptions
            m_powerMenu_onCommand LogOff
        
        
        Case "OPTIONS"
            'frmStartMenuBase.CloseMe
            'frmOptions.Show
            m_powerMenu_onCommand ShowOptions
        
        Case "ABOUT"
            'frmStartMenuBase.CloseMe
            'frmAbout.Show
            m_powerMenu_onCommand ShowAbout
            
        Case "EXIT"
            MainHelper.ExitApplication
            
        Case "HIBERNATE"
            'frmStartMenuBase.CloseMe
            
            PutOptions
            m_powerMenu_onCommand Hibernate
            
        Case "STANDBY"
            'frmStartMenuBase.CloseMe
            
            PutOptions
            m_powerMenu_onCommand StandBy
            
        Case Else
            sP = Split(theItemTag, ";")
            
            For lngIndex = 0 To UBound(sP)
                ShellCommand sP(lngIndex)
            Next
    
    End Select

End Sub

Private Sub m_powerMenu_onInActive()
    m_ignoreActivation = True
    m_arrowButton.AllowUpdates = True
    
Dim cursorPos As win.POINTL

    GetCursorPos cursorPos
    ScreenToClient Me.hWnd, cursorPos
    
    UpdatePowerButtonState m_arrowButton, False, -2, cursorPos.X, cursorPos.Y
    ReDraw
End Sub

Private Sub m_programMenu_onClick(srcNode As INode)
    CloseMe
    
    'If Is64bit() Then
        'ExplorerRun (srcNode.Tag)
    'Else
        'If Not ShellEx(srcNode.Tag) = APITRUE Then Exit Sub
    'End If
    
    SelectBestExecutionMethod srcNode.Tag
    
    Settings.Programs.UpdateByNode srcNode
    m_recentPrograms.PopulateItems
End Sub

Private Sub m_programMenu_onExit(ByVal index As Long)
    Debug.Print "m_programMenu_onExit:: " & index

    g_KeyboardMenuState = 1
    g_KeyboardSide = 2
    
    If index > m_navigationDraw.count Then
        index = m_navigationDraw.count
    End If
    
    ShowRightRollover index

End Sub

Private Sub m_programMenu_onNotifyAllPrograms()
    win.SetFocus Me.hWnd
    
    g_KeyboardMenuState = 1
    m_AllPrograms.Visible = True
    
    ReDraw
End Sub

Private Sub m_programMenu_onRequestCloseStartMenu()
    CloseMe
End Sub

Private Sub m_programMenu_onRequestRecentProgramsRefresh()
    m_recentPrograms.PopulateItems
End Sub

Private Sub m_recentItems_onClickItem(strPath As String)
    CloseMe
    ShellEx strPath
End Sub

Private Sub m_recentItems_onInActive()
    m_RecentItems.Hide
End Sub

Private Sub m_recentPrograms_onRequestCloseStartMenu()
    CloseMe
End Sub

Private Sub m_recentPrograms_onExitSide(ByVal index As Long)
    Debug.Print "m_recentPrograms_onExitSide:: " & index

    g_KeyboardMenuState = 1
    g_KeyboardSide = 2
    
    If index = 0 Then index = 1
    If index > m_navigationDraw.count Then
        index = m_navigationDraw.count
    End If
    
    ShowRightRollover index
End Sub

Private Sub m_recentPrograms_onNotifyAllPrograms()
    g_KeyboardMenuState = 1
    m_AllPrograms.Visible = True
    
    ReDraw
End Sub

Private Sub m_recentPrograms_onRequestShowJumpList(bSuccess As Boolean, theJumpList As JumpList)
    If Not m_JumpListEnabled Then
        MsgBox "The skin doesn't support Jump Lists. Please contact the aurthor of the skin and politely request they update their skin", vbCritical, "Requested resources unavailable"
        Exit Sub
    End If
    
bSuccess = True
    
    If Not m_jumpListMode Then
    
        m_jumpListDrawer.Source = theJumpList
        bSuccess = MorphToJumpList

        Exit Sub
    Else
    
        If m_jumpListDrawer.Source Is theJumpList Then
            MorphToNormal
            Exit Sub
        Else
            m_jumpListDrawer.Source = theJumpList
            
            'bring start menu base to the front every user clicks on a
            'program to show the jumplist contents (this ensures the tooltip is shown)
            Me.Show

            
        End If
    
    End If
    
    ReDraw
End Sub

Private Sub m_searchText_onChange()
    If m_searchingStarted Then
        
        m_research = True
    Else
    
        timTreeViewSearch.Enabled = False
        timTreeViewSearch.Enabled = True
        
    End If
End Sub

Private Sub m_searchText_onFocus()
    If m_searchText.Text = GetPublicString("strStartSearch", "Start Search") Then
        m_searchText.Text = ""
        m_searchText.Font = m_searchBoxNormalFont
        SetTextColor m_searchText.hWnd, vbBlack
        
    End If
End Sub

Private Sub m_searchText_onKeyDown(KeyCode As Long)
    Debug.Print "m_searchText_onKeyDown;;" & KeyCode
    
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        KeyCode = 0
        Exit Sub
    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
               
        If g_KeyboardMenuState = 1 Then
            m_AllPrograms.Visible = False
            g_KeyboardMenuState = 0
            
            ReDraw
        End If
        
        SendKeyToActiveWindow KeyCode
        KeyCode = 0
        
    ElseIf KeyCode = vbKeyReturn Then
    
        If g_KeyboardMenuState <> 1 Then
        
            If AppLauncherHelper.shell32(Me.hWnd, m_searchText.Text) = False Then
                
                If m_recentPrograms.Visible Then
                    m_recentPrograms.RequestExecuteSelected
                Else
                    m_programMenu.RequestActionNode
                End If
                
                KeyCode = 0
            End If
        End If
    
    Else

    End If
End Sub

Private Sub m_searchText_onKeyUp(KeyCode As Long)
    Form_KeyUp CInt(KeyCode), 0
End Sub

Private Sub m_searchText_onLostFocus()
    If m_searchText.Text = vbNullString Then
        m_searchText.Text = GetPublicString("strStartSearch", "Start Search")
        m_searchText.Font = m_searchBoxForeFontItalic
    End If
End Sub

Private Sub m_searchText_onMouseWheel(ByVal wParam As Long)
    PostMessage m_programMenu.hWnd, WM_MOUSEWHEEL, ByVal wParam, 0
End Sub

Private Sub m_userPictureEvents_onClick(ObjSender As Object)
MsgBox ":D"
End Sub

Private Sub m_userPictureEvents_onMouseDown(ObjSender As Object)
MsgBox ":D"
End Sub

Private Sub m_userPictureEvents_onMouseUp(ObjSender As Object, ByVal lngButtonIndex As Long)
    MsgBox ":D"
End Sub

Private Sub m_userPicture_onMouseUp()
    CmdRun "control " & """" & "nusrmgr.cpl" & """"
End Sub

Private Sub timAutoClick_Timer()
    On Error Resume Next
    
    If Not OptionsHelper.bAutoClick Then
        Exit Sub
    End If
    
Dim cursorPos As win.POINTL

    GetCursorPos cursorPos
    ScreenToClient Me.hWnd, cursorPos
    
    Debug.Print cursorPos.X & ":" & cursorPos.Y
    
    m_ignoreActivation = False
    
    Form_MouseDown vbLeftButton, 0, CSng(cursorPos.X), CSng(cursorPos.Y)
    Form_MouseUp vbLeftButton, 0, CSng(cursorPos.X), CSng(cursorPos.Y)
    
    timAutoClick.Enabled = False
End Sub

Private Function UndoLayeredWindow()

    m_layeredMode = False
    m_jumpListMode = False
    
    MoveWindow Me.hWnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, m_background.Width, m_background.Height, False

    SetWindowLong Me.hWnd, GWL_EXSTYLE, _
        GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED And Not WS_EX_TOOLWINDOW
    
    Set m_layeredData = Nothing
    If m_initalized Then ReInitSurface
End Function

Private Function ReBlurMe(UpdateMe As Boolean)

    If Not BlurEnabled Then Exit Function
    
    m_jumpListMode = False
    
    MoveWindow Me.hWnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, m_background.Width, m_background.Height, False

    ReDraw
End Function

Private Function UnblurMe(UpdateMe As Boolean, Optional UndoLayeredWindowOnClose As Boolean = True)
'In order to UnBlur something, we need to make it a layered window

    If m_layeredData Is Nothing Then
        'm_graphics.Clear
        
        Set m_layeredData = MakeLayerdWindow(Me)
        
        m_layeredMode = True
        m_undoLayeredOnClose = UndoLayeredWindowOnClose
        
        SetForegroundWindow Me.hWnd

        m_graphics.FromHDC m_layeredData.theDC
        
        'If UpdateMe Then
            'UpdateBuffer
            'ReDraw
        'End If
    End If
    
    m_graphics.Clear
    
    m_graphics.DrawImage m_navigationDraw.Image, Layout.GroupMenuSchema.Left, Layout.GroupMenuSchema.Top, Layout.GroupMenuSchema.Width, Layout.GroupMenuSchema.Height
    m_graphics.DrawImage m_background, 0, 0, m_background.Width, m_background.Height

    Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, m_layeredData.GetSize, m_layeredData.theDC, m_layeredData.GetPoint, 0, m_blendFunc32bpp, ULW_ALPHA)


End Function

Private Sub MorphStartMenu(ByRef morphStart As GDIPImage, ByRef morphEnd As GDIPImage)

    m_AlphaAmount = 1
    
    Set m_morphFrom = morphStart
    Set m_morphTo = morphEnd

    timMorphToJumpList.Enabled = True
End Sub

Private Sub timJumplistUpdater_Timer()
    If isset(g_AutomaticDestinationsUpdater) Then
        g_AutomaticDestinationsUpdater.Update
    End If
End Sub

Private Sub timMorphToJumpList_Timer()

    m_AlphaAmount = m_AlphaAmount - 0.1
    'TODO: Update GDIPlus library to allow the below code to work
    
'    m_theMatrix.m(3, 3) = m_AlphaAmount
'    m_theMatrix2.m(3, 3) = 1 - m_AlphaAmount
    

'Dim imageAttrib As New GDIPImageAttributes
'Dim imageAttrib2 As New GDIPImageAttributes

'    imageAttrib.SetColorMatrix m_theMatrix
'    imageAttrib2.SetColorMatrix m_theMatrix2
    
'    m_graphics.Clear
    
'    If m_morphFrom Is m_background Then m_graphics.DrawImage m_navigationDraw.Image, Layout.GroupMenuSchema.Left, Layout.GroupMenuSchema.Top, Layout.GroupMenuSchema.Width, Layout.GroupMenuSchema.Height, imageAttrib
'    m_graphics.DrawImage _
            m_morphFrom, 0, 0, m_morphFrom.Width, m_morphFrom.Height, imageAttrib


 '   If m_morphTo Is m_background Then m_graphics.DrawImage m_navigationDraw.Image, Layout.GroupMenuSchema.Left, Layout.GroupMenuSchema.Top, Layout.GroupMenuSchema.Width, Layout.GroupMenuSchema.Height, imageAttrib2
 '   m_graphics.DrawImage m_morphTo, 0, 0, m_morphTo.Width, m_morphTo.Height, imageAttrib2

    
 '   If Not m_layeredData Is Nothing Then
 '       Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, m_layeredData.GetSize, m_layeredData.theDC, m_layeredData.GetPoint, 0, m_blendFunc32bpp, ULW_ALPHA)
 '   End If

    If m_AlphaAmount <= 0 Then
        timMorphToJumpList.Enabled = False

        ReDraw
        
        Me.Show
        
        If m_reBlurcallback Then
            ReBlurMe False
        End If
    End If
        
End Sub

Private Sub timFade_Timer()
On Error Resume Next
'Dim cursorPos As win.POINTL

    'GetCursorPos cursorPos
    'ScreenToClient Me.hWnd, cursorPos
    
    'Form_MouseMove GetAsyncKeyState(User.VK_LBUTTON), 0, CSng(cursorPos.x), CSng(cursorPos.y)

    If Not m_currentFadeInImage Is Nothing Then
        If m_currentFadeInImage.Alpha < 255 Then
            m_currentFadeInImage.Alpha = m_currentFadeInImage.Alpha + 15
        End If
    End If
    
Dim thisImage As frmRolloverImage
    For Each thisImage In m_FadeOuts
        If Not thisImage Is Nothing Then
            If thisImage.Alpha > 0 Then
                thisImage.Alpha = thisImage.Alpha - 15
            Else
                m_FadeOuts.Remove thisImage.Tag
                Unload thisImage
                
            End If
        End If

    Next

End Sub

Private Sub timRolloverDelay_Timer()
    If FileExists(m_RolloverToShow) Then
        ShowRollover m_RolloverToShow
    End If
End Sub

Private Sub timTreeViewSearch_Timer()
    m_searchingStarted = True

Dim searchText As String
    searchText = m_searchText.Text
    
    g_KeyboardMenuState = 0
    g_KeyboardSide = 1
    
    m_AllPrograms.Visible = False
    m_navigationDraw.ResetRollover
    
    ReDraw
    
    If searchText <> GetPublicString("strStartSearch", "Start Search") And _
        searchText <> vbNullString Then
        
        If m_recentPrograms.Visible Then
            ToggleProgramsMenu
            Debug.Print "Toggle!"
        End If
            
        m_programMenu.Filter = searchText
    Else
        m_programMenu.Filter = vbNullString
    End If
    
    'm_programMenu.SelectFirstVisibleItem
    
    timTreeViewSearch.Enabled = False
    m_searchingStarted = False
    
    If m_research = True Then
        m_research = False
        
        timTreeViewSearch_Timer
    End If
End Sub

Private Function GenerateRolloverImage(Path As String)

Dim rolloverImage As New frmRolloverImage
Static rolloverID As Long
    
    rolloverImage.Path = Path
    rolloverImage.MakeTrans Path

    'While rolloverImage.Tag = vbNullString
    rolloverImage.Tag = rolloverID
    
    'probably impossible
    If rolloverID = 2147483468 Then
        rolloverID = 0
    End If
    rolloverID = rolloverID + 1
    'Wend
    
    Set GenerateRolloverImage = rolloverImage
End Function

Private Function UpdateAutoClickButton(cursorPosition As POINTL, sourceRect As RECT)

    If cursorPosition.X > sourceRect.Left And _
        cursorPosition.Y > sourceRect.Top And _
         cursorPosition.Y < sourceRect.Bottom And _
          cursorPosition.X < sourceRect.Right Then
          
          timAutoClick.Enabled = True
    End If

End Function

