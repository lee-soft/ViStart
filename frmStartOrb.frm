VERSION 5.00
Begin VB.Form frmStartOrb 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ViStart_PngNew"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   21
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmStartOrb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LAYERD_MODE As Long = 1

Dim srcPoint As POINTL

Dim Snaps(3) As gdiplus.RECTL
Dim bMouseOn As Boolean

Private m_pngHeight As Long
Private m_pngWidth As Long
Private m_pngPath As String

' Force the callback class to implement the interface defining the events.
Private mCallback As IPngImageEvents

Private m_graphicsImage As GDIPGraphics
Private m_imageOrb As GDIPImage
Private m_OrbLayout As gdiplus.RECTL
Private m_win78 As GDIPImage

Private m_lastDrawnSnap As Long
Private m_heightDifference As Long
Private m_handles As LayerdWindowHandles

Private WithEvents m_startOptions As frmVistaMenu
Attribute m_startOptions.VB_VarHelpID = -1
Private WithEvents m_changeSkin As frmSkinSelect
Attribute m_changeSkin.VB_VarHelpID = -1
Private WithEvents m_optionsDialog As frmControlPanel
Attribute m_optionsDialog.VB_VarHelpID = -1

Private m_max_height As Long

Private m_blendMode As BLENDFUNCTION

Private m_mode As Long
Private m_buttonDownCount As Long

Public FullHeight As Boolean
Public Windows81LeftDifference As Long

Implements IHookSink

Private m_logger As SeverityLogger

Property Get Logger()
    Set Logger = m_logger
End Property

Public Function SetContextMenu(ByRef newContextMenu As frmVistaMenu)
    Set m_startOptions = newContextMenu
End Function

Public Function ResetOrb()

        If FileExists(ResourcesPath & "start_button.png") Then
                Me.Path = ResourcesPath & "start_button.png"
        ElseIf FileExists(sCon_AppDataPath & "_orbs\" & Settings.CurrentOrb) Then
                Me.Path = sCon_AppDataPath & "_orbs\" & Settings.CurrentOrb
        ElseIf FileExists(sCon_AppDataPath & "_orbs\default.png") Then
                Me.Path = sCon_AppDataPath & "_orbs\default.png"
        ElseIf FileExists(sCon_AppDataPath & "_orbs\Windows 7.png") Then
                Me.Path = sCon_AppDataPath & "_orbs\Windows 7.png"
        ElseIf FileExists(sCon_AppDataPath & "_orbs\start_button.png") Then
                Me.Path = sCon_AppDataPath & "_orbs\start_button.png"
        ElseIf FileExists(sCon_AppDataPath & "start_button.png") Then
                Me.Path = sCon_AppDataPath & "start_button.png"
        End If

    Settings.CurrentOrb = vbNullString
    
End Function

Public Property Get MaxHeight() As Long
    MaxHeight = m_max_height
End Property

Public Property Let MaxHeight(newHeight As Long)
    If Not m_mode = LAYERD_MODE Then
        m_max_height = newHeight
        Test2
        
        DrawPart 0
    End If
    

End Property

Public Property Get mode() As Long
    mode = m_mode
End Property

Public Property Let mode(newMode As Long)
    m_mode = newMode
    
    If newMode = LAYERD_MODE Then
        Me.AutoRedraw = False
        
        With m_blendMode
           .AlphaFormat = AC_SRC_ALPHA ' 32 bit
           .BlendFlags = 0
           .BlendOp = AC_SRC_OVER
           .SourceConstantAlpha = 255
        End With
    End If
End Property

Function DrawImageStretchRect(ByRef Image As GDIPImage, ByRef destRect As gdiplus.RECTL, ByRef sourceRect As gdiplus.RECTL)
    m_graphicsImage.DrawImageStretchAttrL Image, _
        destRect, _
        sourceRect.Left, sourceRect.Top, sourceRect.Width, sourceRect.Height, UnitPixel, 0, 0, 0
End Function

Property Let Path(ByVal strPath As String)
    If m_mode = 0 Then
        Logger.Fatal "Mode Not Set!", "Path", strPath
        Exit Property
    End If

    m_pngPath = strPath
    
    Set m_graphicsImage = New GDIPGraphics
    Set m_imageOrb = New GDIPImage
    
    m_imageOrb.Dispose
    m_imageOrb.FromFile strPath
    
    If m_imageOrb.Width > 50 Then
        Windows81LeftDifference = 50 - m_imageOrb.Width
    End If
    
    Me.Width = m_imageOrb.Width * Screen.TwipsPerPixelX
    Me.Height = ((m_imageOrb.Height / 3) * Screen.TwipsPerPixelY)
    
    
    m_heightDifference = 0
    
    m_OrbLayout.Top = 0
    m_OrbLayout.Left = 0
    
    m_OrbLayout.Width = Me.ScaleWidth
    m_OrbLayout.Height = Me.ScaleHeight
    
    SetSnap 0, 0, m_imageOrb.Width, 0, Me.ScaleHeight
    SetSnap 1, 0, m_imageOrb.Width, Me.ScaleHeight, Me.ScaleHeight
    SetSnap 2, 0, m_imageOrb.Width, Me.ScaleHeight * 2, Me.ScaleHeight
    
    If Not MainHelper.g_viOrb_fullHeight Then
        m_OrbLayout.Top = m_OrbLayout.Top - 1
    End If
    
    If Not m_mode = LAYERD_MODE Then
        If (Me.ScaleHeight > m_max_height) Then
            m_heightDifference = Me.ScaleHeight - m_max_height
            Me.Height = m_max_height * Screen.TwipsPerPixelY
            
            m_OrbLayout.Top = -(m_heightDifference / 2)
        End If
        
        m_graphicsImage.FromHDC Me.hdc
    End If
    
    If m_mode = LAYERD_MODE Then
        Set m_handles = MakeLayerdWindow(Me)
        m_graphicsImage.FromHDC m_handles.theDC
    End If
    
    m_max_height = 0
    DrawPart 0
End Property

Property Get Path() As String
    Path = m_pngPath
End Property

' Allow the callback object to be set. Very important.
Property Set callback(ByRef newObj As IPngImageEvents)
    Set mCallback = newObj
End Property

Property Get callback() As IPngImageEvents
    Set callback = mCallback
End Property

Public Sub SetSnap(index As Long, lLeft As Long, lRight As Long, lTop As Long, lBottom As Long)

    With Snaps(index)
        .Height = lBottom
        .Left = lLeft
        .Width = lRight
        .Top = lTop
    End With

End Sub

Private Sub Form_Click()

    ' Raise an event, passing a parameter
    If (Not mCallback Is Nothing) Then _
        mCallback.onClick Me
        
End Sub

Private Sub Form_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)

    Set m_optionsDialog = frmControlPanel
    bMouseOn = True

    If ShellHelper.g_Windows7 Or ShellHelper.g_Windows8 Then
        Set m_win78 = New GDIPImage
        
        If FileExists(ResourcesPath & "orb_background.png") Then
            m_win78.FromFile ResourcesPath & "orb_background.png"
        Else
            m_win78.FromBinary LoadResData("WIN7", "PNG")
        End If
        
    End If
    
    DragAcceptFiles Me.hWnd, APITRUE
    HookWindow Me.hWnd, Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim flagSkipMetro As Boolean

    If Button = vbRightButton Then

        If g_bStartMenuVisible Then
            'Make StartMenu be inactivate
            PostMessage frmStartMenuBase.hWnd, WM_ACTIVATEAPP, ByVal MakeLong(0, WA_INACTIVE), 0
        End If
        m_startOptions.Resurrect True, frmEvents
    
    ElseIf Button = vbLeftButton Then
        DrawPart 2
    
        ' Raise an event, passing a parameter
        If (Not mCallback Is Nothing) Then _
            mCallback.onMouseDown Me
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Raise an event, passing a parameter
    If (Not mCallback Is Nothing) Then _
        mCallback.onMouseUp Me, Button

End Sub

Sub DrawPart(iSnapIndex As Integer)
    If m_graphicsImage Is Nothing Then Exit Sub

    m_lastDrawnSnap = iSnapIndex
    m_graphicsImage.Clear
    
    If Not mode = LAYERD_MODE Then
        m_graphicsImage.DrawImage m_win78, 0, 0, Me.ScaleWidth * 2, Me.ScaleHeight * 2
    End If
    
    DrawImageStretchRect m_imageOrb, m_OrbLayout, Snaps(iSnapIndex)
    
    If m_mode = LAYERD_MODE Then
        Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, m_handles.GetSize, m_handles.theDC, m_handles.GetPoint, 0, m_blendMode, ULW_ALPHA)
    Else
        Me.Refresh
    End If
End Sub

Sub Test2()

    If Not mode = LAYERD_MODE Then
        Me.Width = m_imageOrb.Width * Screen.TwipsPerPixelX
        Me.Height = m_max_height * Screen.TwipsPerPixelY
        
        m_heightDifference = 0
        
        m_OrbLayout.Width = Me.ScaleWidth
        m_OrbLayout.Height = (m_imageOrb.Height / 3)
        
        m_OrbLayout.Top = (m_max_height / 2) - (m_OrbLayout.Height / 2)
        If Not MainHelper.g_viOrb_fullHeight Then
            m_OrbLayout.Top = m_OrbLayout.Top - 1
        End If
        
        m_OrbLayout.Left = 0
    
        SetSnap 0, 0, m_imageOrb.Width, 0, m_OrbLayout.Height
        SetSnap 1, 0, m_imageOrb.Width, m_OrbLayout.Height, m_OrbLayout.Height
        SetSnap 2, 0, m_imageOrb.Width, m_OrbLayout.Height * 2, m_OrbLayout.Height
    
        If (Me.ScaleHeight > m_max_height) Then
            m_heightDifference = Me.ScaleHeight - m_max_height
            Me.Height = m_max_height * Screen.TwipsPerPixelY
            
            m_OrbLayout.Top = -(m_heightDifference / 2)
        End If
        
        m_graphicsImage.FromHDC Me.hdc
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_handles = Nothing
    UnhookWindow Me.hWnd
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
    On Error GoTo Handler
    
    If msg = WM_DROPFILES Then

        Dim hFilesInfo As Long
        Dim szFileName As String
        Dim wTotalFiles As Long
        Dim wIndex As Long
        
        hFilesInfo = wp
        wTotalFiles = DragQueryFileW(hFilesInfo, &HFFFF, ByVal 0&, 0)
    
        For wIndex = 0 To wTotalFiles
            szFileName = Space$(1024)
            
            If Not DragQueryFileW(hFilesInfo, wIndex, StrPtr(szFileName), Len(szFileName)) = 0 Then
                Form_DragDropFile TrimNull(szFileName)
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
    Logger.Error Err.Description, "IHookSink_WindowProc"

    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
        CallOldWindowProcessor(hWnd, msg, wp, lp)
End Function

Private Sub m_changeSkin_onChangeSkin(szNewSkin As String)
    frmStartMenuBase.Skin = szNewSkin
End Sub

Private Sub m_optionsDialog_onChangeOrb(szNewOrb As String)
    If szNewOrb = vbNullString Then
        Me.ResetOrb
    Else
        Me.Path = sCon_AppDataPath & "_orbs\" & szNewOrb
        Settings.CurrentOrb = szNewOrb
    End If
End Sub

Public Function PromptUserSelectNewOrb() As String

Dim newStartPng As String
Dim newStartPngName As String

    newStartPng = BrowseForFile(sCon_OrbFolderPath, "PNG (*.PNG);*.png", GetPublicString("strViStartOrb"), frmEvents.hWnd)

    If InstallOrb(newStartPng, newStartPngName) = True Then
    
        Settings.CurrentOrb = newStartPngName
        Me.Path = sCon_AppDataPath & "_orbs\" & newStartPngName
    End If

End Function

Public Function Form_DragDropFile(ByVal szFileName As String)

Dim recentPrograms As frmFreq

    Settings.Programs.TogglePin_ElseAddToPin_ByProgram CreateProgramFromPath(szFileName)
End Function
