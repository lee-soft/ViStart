VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViNavigationPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NavigationPaneOptions As Collection
Option Explicit

Public Event onChanged()
Public Event onCommand(szCommand As String)

Public HasMouse As Boolean

Private Const TOP_MARGIN As Long = 10
Private Const LEFT_MARGIN As Long = 7
Private Const ROLLOVER_X_MARGIN As Long = 2

Private m_navigationPaneOptions As Collection
Private m_visiblePaneOptions As Collection

Private m_graphics As GDIPGraphics
Private m_surface As GDIPBitmap

Private m_background As GDIPImage
Private m_backgroundPosition As POINTL

Private m_Width As Long
Private m_Height As Long

Private m_font As GDIPFont
Private m_Brush As GDIPBrush

Private m_seperatorAmount As Long
Private m_selectedItem As Object
Private m_itemToBeRenamed As Object

Private m_rollover As GDIPImage
Private m_arrow As GDIPImage

Private m_rolloverHeight As Long
Private m_lastIndex As Long
Private m_mouseDown As Boolean
Private m_rolloverPosition As POINTL

Private WithEvents m_contextMenu As frmVistaMenu
Attribute m_contextMenu.VB_VarHelpID = -1
Private WithEvents m_customContextMenu As frmVistaMenu
Attribute m_customContextMenu.VB_VarHelpID = -1

Private WithEvents m_popoutMenu As frmFileMenu
Attribute m_popoutMenu.VB_VarHelpID = -1

Private m_windowPosition As POINTL
Private m_window As Form

Private m_RegObj As New RegistryKey

Private GroupOptionsVisibileCount As Long
Private MinOSVersion As Double
Private MaxOSVersion As Double

Private Declare Function SHChangeNotify Lib "shell32.dll" (ByVal wEventID As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long) As Long

Private WithEvents m_renameTextBox As VistaSearchBox
Attribute m_renameTextBox.VB_VarHelpID = -1

Private m_logger As SeverityLogger

Property Get Logger() As SeverityLogger
    Set Logger = m_logger
End Property

Public Property Get NavigationOptions() As Collection
    Set NavigationOptions = m_navigationPaneOptions
End Property

Public Function AddNewFolder(szNewFolderPath As String)
    
    Dim thisFolderOption As New NavigationPaneFolder
    Dim thisFolder As Scripting.Folder

    If FSO.FolderExists(szNewFolderPath) = False Then
        Exit Function
    End If
    
    Set thisFolder = FSO.GetFolder(szNewFolderPath)
    
    thisFolderOption.OpenAsMenu = False
    thisFolderOption.Caption = thisFolder.name
    thisFolderOption.Shell = thisFolder.Path
    thisFolderOption.Visible = True
    
    If thisFolderOption.Caption = "" Then thisFolderOption.Caption = thisFolderOption.Shell
    
    m_navigationPaneOptions.Add thisFolderOption
    PositionOptions
    
    RaiseEvent onChanged
End Function

Public Function RenameWindowHasFocus() As Boolean
    If IsWindowVisible(m_renameTextBox.hWnd) = APITRUE Then
        RenameWindowHasFocus = True
    End If
End Function

Public Function ShowContextMenu(ByRef winOwner As Form)

    On Error GoTo Handler

    If (m_selectedItem Is Nothing) Then
        ShowContextMenu = False
        Exit Function
    End If
    
    If Not m_contextMenu Is Nothing Then Unload m_contextMenu
    If Not m_customContextMenu Is Nothing Then m_customContextMenu.Hide

    Dim DesktopIconStatus As Variant
    Dim ComputerIconStatus As Variant
    
    Dim NewStartPanelRegKey As RegistryKey
    Set NewStartPanelRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel")
    If NewStartPanelRegKey Is Nothing Then
        Set NewStartPanelRegKey = Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel")
    End If
    
    Dim HideMyComputerIconsRegKey As RegistryKey
    Set HideMyComputerIconsRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideMyComputerIcons")
    If HideMyComputerIconsRegKey Is Nothing Then
        Set HideMyComputerIconsRegKey = Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideMyComputerIcons")
    End If
        
    If TypeName(m_selectedItem) = "NavigationPaneFolder" Then

        Set m_contextMenu = New frmVistaMenu
    
        m_contextMenu.AddItem GetPublicString("strOpen"), "OPEN_OPTION", True

        m_contextMenu.AddItem ""
        If m_selectedItem.OpenAsMenu Then
            m_contextMenu.AddItem GetPublicString("strDisplayAsLink"), "DONT_SHOW_MENU", False
        Else
            m_contextMenu.AddItem GetPublicString("strDisplayAsMenu"), "SHOW_MENU", False
        End If

        m_contextMenu.AddItem ""
        m_contextMenu.AddItem GetPublicString("strDontShowItem"), "HIDE_OPTION", False
        		
                DesktopIconStatus = 0
                ComputerIconStatus = 0
                
                If InStr(1, UCase(m_selectedItem.Shell), g_CLSID_USERPROFILE, 1) > 0 Then
                        ' %UserProfile%
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_USERPROFILE)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_USERPROFILE)
                                
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_LIBRARIES, 1) > 0 Then
                        ' Libraries
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_LIBRARIES)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_LIBRARIES)
                
                ElseIf InStr(1, UCase(m_selectedItem.Shell), "{450D8FBA-AD25-11D0-98A8-0800361B1103}", 1) > 0 Or InStr(1, UCase(m_selectedItem.Shell), "PERSONAL", 1) Or InStr(1, UCase(m_selectedItem.Shell), g_CLSID_MYDOCS, 1) > 0 Then
                        ' My documents
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYDOCS)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYDOCS)
                        
                        If DesktopIconStatus = 1 Then
                                ' Documents
                                DesktopIconStatus = NewStartPanelRegKey.GetValue("{450D8FBA-AD25-11D0-98A8-0800361B1103}")
                        End If
                        If ComputerIconStatus = 1 Then
                                ' Documents
                                ComputerIconStatus = HideMyComputerIconsRegKey.GetValue("{450D8FBA-AD25-11D0-98A8-0800361B1103}")
                        End If

                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_MYPIC, 1) > 0 Or InStr(1, UCase(m_selectedItem.Shell), "MYPICTURES", 1) Then
                        ' Pictures
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYPIC)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYPIC)
                                                
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_MYMUS, 1) > 0 Or InStr(1, UCase(m_selectedItem.Shell), "MYMUSIC", 1) Then
                        ' Music
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYMUS)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYMUS)
                        
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_MYVID, 1) > 0 Or InStr(1, UCase(m_selectedItem.Shell), "MYVIDEO", 1) Then
                        ' Videos
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYVID)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYVID)
                                                
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_DOWNLOADS, 1) > 0 Then
                        ' Downloads
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_DOWNLOADS)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_DOWNLOADS)
                           
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_3DOBJECTS, 1) > 0 Then
                        ' 3D Objects
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_3DOBJECTS)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_3DOBJECTS)
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_DESKTOP, 1) > 0 Then
                        ' Desktop
                
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_DESKTOP)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_DESKTOP)
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_GAMES, 1) > 0 Then
                        ' Games Explorer
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_GAMES)
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_GAMES)
                Else
                        DesktopIconStatus = -2
                        ComputerIconStatus = -2
                End If
                
                ' -1 REG key doesn't exist
                ' -2 Other items
                
				If DesktopIconStatus = 1 Or ComputerIconStatus = 1 Then
					m_contextMenu.AddItem ""
				End If
				
                If DesktopIconStatus = 1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strShowOnDesktop"), "SHOWHIDE_D"
                ElseIf DesktopIconStatus = 0 Or DesktopIconStatus = -1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strHideFromDesktop"), "SHOWHIDE_D"
                End If
                
                If ComputerIconStatus = 1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strShowInComputer"), "SHOWHIDE_C"
                ElseIf ComputerIconStatus = 0 Or ComputerIconStatus = -1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strHideFromComputer"), "SHOWHIDE_C"
                End If



            
        m_contextMenu.AddItem ""
        m_contextMenu.AddItem GetPublicString("strRename"), "RENAME_OPTION"
        m_contextMenu.AddItem ""
        m_contextMenu.AddItem GetPublicString("strProperties"), "PROPERTIES_OPTION"
    
        m_contextMenu.Resurrect True, winOwner
    
    Else
    
        If m_selectedItem.OptionMenu.CountItems > 0 Then
            Set m_customContextMenu = m_selectedItem.OptionMenu
            m_customContextMenu.Resurrect True, winOwner
        End If
    
                Set m_contextMenu = New frmVistaMenu
        

                If Left(UCase(m_selectedItem.Shell), 8) = ":FILERUN" Then
                        m_contextMenu.AddItem GetPublicString("strRun"), "RUN", True
                Else
                    m_contextMenu.AddItem GetPublicString("strOpen"), "EXECUTE", True
                End If
                
                DesktopIconStatus = 0
                ComputerIconStatus = 0

                If InStr(1, UCase(m_selectedItem.Shell), g_CLSID_COMPUTER, 1) > 0 Then
                        ' This PC
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_COMPUTER)
                        ComputerIconStatus = -2
                        
                ElseIf InStr(1, UCase(m_selectedItem.Shell), g_CLSID_NETWORK, 1) > 0 Then
                        ' My network places
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_NETWORK)
                        ComputerIconStatus = -2
                        
                ElseIf Left(UCase(m_selectedItem.Shell), 7) = "CONTROL" Then
                                ' Control Panel
                        If WindowsVersion < 6 Then
                                If FSO.FolderExists(CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(CreateObject("WScript.Shell").SpecialFolders("Desktop")) & "\" & GetPublicString("strControlPanel") & "." & g_CLSID_CONTROLPANEL) Then
                                DesktopIconStatus = 0
                        Else
                                DesktopIconStatus = 1
                        End If
                Else
                                DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_CONTROLPANEL)
                                ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_CONTROLPANEL)
                        End If
                Else
                        DesktopIconStatus = -2
                        ComputerIconStatus = -2
                End If
                
                ' -1 REG key doesn't exist
                ' -2 Other items
                
				If DesktopIconStatus = 1 Or ComputerIconStatus = 1 Then
					m_contextMenu.AddItem ""
				End If
				
                If DesktopIconStatus = 1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strShowOnDesktop"), "SHOWHIDE_D"
                ElseIf DesktopIconStatus = 0 Or DesktopIconStatus = -1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strHideFromDesktop"), "SHOWHIDE_D"
                End If
                
                If ComputerIconStatus = 1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strShowInComputer"), "SHOWHIDE_C"
                ElseIf ComputerIconStatus = 0 Or ComputerIconStatus = -1 Then
                        'm_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strHideFromComputer"), "SHOWHIDE_C"
                End If
                
        m_contextMenu.AddItem ""
        m_contextMenu.AddItem GetPublicString("strDontShowItem"), "HIDE_OPTION", False

        m_contextMenu.AddItem ""
                
                If InStr(1, UCase(m_selectedItem.Shell), g_CLSID_COMPUTER, 1) > 0 Then
                        ' This PC
                        m_contextMenu.AddItem GetPublicString("strSearch"), "SEARCH"
                        m_contextMenu.AddItem GetPublicString("strManage"), "MANAGE"
                        m_contextMenu.AddItem ""

                End If
                
        m_contextMenu.AddItem GetPublicString("strRename"), "RENAME_OPTION"
        
                If Left(UCase(m_selectedItem.Shell), 12) = "EXPLORER.EXE" Then
                
                        m_contextMenu.AddItem ""
                        m_contextMenu.AddItem GetPublicString("strProperties"), "SYSPROPERTIES"

                End If
                
                m_contextMenu.Resurrect True, winOwner
                
        End If

    Exit Function
Handler:
    Logger.Error Err.Description, "ShowContextMenu"

End Function

Public Function ExecutedSelected()
    If Not m_selectedItem Is Nothing Then
        RaiseEvent onCommand(m_selectedItem.Shell)
    End If
End Function

Public Function count() As Long
    count = m_visiblePaneOptions.count
End Function

Public Function ResetRollover()
    m_lastIndex = -1
    Set m_selectedItem = Nothing
    
    RaiseEvent onChanged
End Function

Public Property Get index() As Long
    index = m_lastIndex
End Property

Public Function SelectOptionByIndex(newIndex As Long)
    If newIndex > 0 And newIndex <= m_visiblePaneOptions.count Then
        m_lastIndex = newIndex
        Set m_selectedItem = m_visiblePaneOptions(m_lastIndex)
        MoveRolloverToIndex m_lastIndex

        RaiseEvent onChanged
    Else
        Logger.Warn newIndex & " index is out of bounds!!", "SelectOptionByIndex"
    End If

End Function

Public Function DumpOptions(ByRef sourceDoc As DOMDocument, ByRef parentElement As IXMLDOMElement) As Boolean
    If m_navigationPaneOptions Is Nothing Then Exit Function

Dim XML_navigationPane As IXMLDOMElement

Dim thisItem As Object
Dim newItem As IXMLDOMElement

    'Set m_sourceDoc = New DOMDocument
    
    Set XML_navigationPane = sourceDoc.createElement("navigation_pane")
    parentElement.appendChild XML_navigationPane
    
    For Each thisItem In m_navigationPaneOptions
        Select Case TypeName(thisItem)
        
        Case "NavigationPaneFolder"
            Set newItem = sourceDoc.createElement("folder")
            XML_navigationPane.appendChild newItem
            
            newItem.setAttribute "caption", thisItem.Caption
            newItem.setAttribute "path", thisItem.Shell
            newItem.setAttribute "rollover", thisItem.Rollover
            newItem.setAttribute "display_as_menu", CStr(thisItem.OpenAsMenu)
            newItem.setAttribute "visible", CStr(thisItem.Visible)
            
        Case "NavigationPaneCustom"
            Set newItem = sourceDoc.createElement("custom")
            XML_navigationPane.appendChild newItem
            
            newItem.setAttribute "caption", thisItem.Caption
            newItem.setAttribute "exec", thisItem.Shell
            newItem.setAttribute "rollover", thisItem.Rollover
            newItem.setAttribute "visible", CStr(thisItem.Visible)
            
            thisItem.DumpContextMenu sourceDoc, newItem
            
        Case Else
            MsgBox "'" & TypeName(thisItem) & "' was unexpected!", vbCritical
            
        End Select
    Next

    DumpOptions = True
End Function

Public Function SelectOptionAbove()
    If m_lastIndex - 1 >= 1 Then
        m_lastIndex = m_lastIndex - 1
    Else
        If m_visiblePaneOptions.count > 0 Then
            m_lastIndex = m_visiblePaneOptions.count
        End If
    End If

    If m_lastIndex > 0 And m_lastIndex <= m_visiblePaneOptions.count Then
        MoveRolloverToIndex m_lastIndex
    Else
        m_lastIndex = m_visiblePaneOptions.count
    End If
End Function

Public Function SelectOptionBelow()
    If m_lastIndex + 1 <= m_visiblePaneOptions.count Then
        m_lastIndex = m_lastIndex + 1
    Else
        If m_visiblePaneOptions.count > 0 Then
            m_lastIndex = 1
        End If
    End If
    
    If m_lastIndex > 0 And m_lastIndex <= m_visiblePaneOptions.count Then
        MoveRolloverToIndex m_lastIndex
    Else
        m_lastIndex = m_visiblePaneOptions.count
    End If
End Function

Public Property Get SelectedItem() As Object
    Set SelectedItem = m_selectedItem
End Property

Public Property Let BaseWindow(newWindow As Form)
    Set m_window = newWindow
    SetOwner m_renameTextBox.hWnd, m_window.hWnd
End Property

Public Property Get LeftMouseButtonDown() As Boolean
    LeftMouseButtonDown = m_mouseDown
End Property

Public Function Constructor(szResourcesPath As String, ByRef layoutFile As LayoutParser, ByRef newBackground As GDIPImage, ByRef navigationPaneXML As IXMLDOMElement)

    m_Width = layoutFile.GroupMenuSchema.Width
    m_Height = layoutFile.GroupMenuSchema.Height
    
    If m_Width = 0 Or m_Height = 0 Then
        Constructor = False
        Exit Function
    End If

    Set m_navigationPaneOptions = New Collection
    
    Set m_background = newBackground
    m_backgroundPosition.X = Layout.GroupMenuSchema.Left
    m_backgroundPosition.Y = Layout.GroupMenuSchema.Top
    
    Set m_rollover = New GDIPImage
    m_rollover.FromFile szResourcesPath & "button.png"
    
    m_rolloverHeight = m_rollover.Height / 2
    
    m_seperatorAmount = 35
    If Layout.GroupOptionsSeparator <> 0 Then
        m_seperatorAmount = layoutFile.GroupOptionsSeparator
    End If

    pInitialize
    
    PopulateLinks navigationPaneXML
    PositionOptions
End Function

Public Function MouseOut()
    If IsContextMenusOpen Then Exit Function
    
    If Not m_selectedItem Is Nothing Then
    
        Set m_selectedItem = Nothing
        m_lastIndex = -1
        RaiseEvent onChanged
    End If
    
End Function

Public Function MouseUp()
    
    If m_mouseDown Then
        m_mouseDown = False
        RaiseEvent onChanged
    End If
End Function

Public Function MouseLeftClick()
    If IsContextMenusOpen Then Exit Function
    
    If Not m_mouseDown Then
        m_mouseDown = True
        RaiseEvent onChanged
        
        SelectItem
    End If
End Function

Private Function ShowFileMenu()
    m_windowPosition.Y = m_window.Top / Screen.TwipsPerPixelY
    m_windowPosition.X = m_window.Left / Screen.TwipsPerPixelX

Dim windowPositonX As Long
Dim windowPositonY As Long

    If Not m_popoutMenu Is Nothing Then Unload m_popoutMenu
    Set m_popoutMenu = New frmFileMenu
    
    m_popoutMenu.PopulateFromPath VarScan(m_selectedItem.Shell)

    windowPositonY = m_windowPosition.Y - m_rolloverHeight
    
    'MsgBox (m_windowPosition.Y + m_backgroundPosition.Y + (m_rolloverPosition.Y + m_rolloverHeight)) - m_popoutMenu.ScaleHeight
    
    If (m_windowPosition.Y + m_backgroundPosition.Y + m_rolloverPosition.Y + m_rolloverHeight + m_popoutMenu.ScaleHeight) > Screen.Height / Screen.TwipsPerPixelY Then
        windowPositonY = (m_windowPosition.Y + m_backgroundPosition.Y + (m_rolloverPosition.Y + m_rolloverHeight)) - m_popoutMenu.ScaleHeight
    Else
        windowPositonY = m_windowPosition.Y + m_backgroundPosition.Y + m_rolloverPosition.Y
    End If
    
    If (m_windowPosition.X + m_backgroundPosition.X + m_rolloverPosition.X + m_rollover.Width + m_popoutMenu.ScaleWidth) > Screen.Width / Screen.TwipsPerPixelX Then
        windowPositonX = (m_windowPosition.X + m_backgroundPosition.X) - (m_popoutMenu.ScaleWidth)
    Else
        windowPositonX = m_windowPosition.X + m_backgroundPosition.X + m_rolloverPosition.X + m_rollover.Width
    End If
    
    'MsgBox m_rolloverPosition.Y
    
    MoveWindow m_popoutMenu.hWnd, windowPositonX, windowPositonY, m_popoutMenu.ScaleWidth, m_popoutMenu.ScaleHeight, True
    SetOwner m_popoutMenu.hWnd, m_window.hWnd
    m_popoutMenu.Show
End Function

Public Function SelectItem()
    
    If Not m_selectedItem Is Nothing Then
        Select Case TypeName(m_selectedItem)
        
        Case "NavigationPaneFolder"
            If m_selectedItem.OpenAsMenu Then
                ShowFileMenu
            Else
                RaiseEvent onCommand("explorer " & """" & VarScan(m_selectedItem.Shell) & """")
            End If
            
        Case "NavigationPaneCustom"
            RaiseEvent onCommand(m_selectedItem.Shell)
            
        Case Else
            MsgBox TypeName(m_selectedItem)
            
        End Select
    End If
    
    
End Function

Private Function MoveRolloverToIndex(newIndex As Long)

    Set m_selectedItem = m_visiblePaneOptions(newIndex)
    
    m_rolloverPosition.Y = m_selectedItem.Position - (m_rolloverHeight / 2) + (m_font.Size / 2) + ROLLOVER_X_MARGIN
    m_rolloverPosition.X = 0
    
    RaiseEvent onChanged

End Function

Public Function MouseMove(Position As POINTL)
    If m_visiblePaneOptions Is Nothing Then Exit Function
    If IsContextMenusOpen Then Exit Function
    If RenameWindowHasFocus Then Exit Function

Dim suggestedIndex As Long
Dim A As Long
Dim newY As Long

    newY = Position.Y - TOP_MARGIN + (m_rolloverHeight / 2) - (m_font.Size / 2) - 2
    suggestedIndex = FindIndex(newY, m_seperatorAmount) + 1

    If suggestedIndex > 0 And suggestedIndex <= m_visiblePaneOptions.count Then
        
        
        A = ((suggestedIndex - 1) * m_seperatorAmount) _
                + m_rolloverHeight
        
        If Position.Y <= A Then
            If m_lastIndex <> suggestedIndex Then
                m_lastIndex = suggestedIndex
                MoveRolloverToIndex suggestedIndex
            End If
        Else
            If Not m_selectedItem Is Nothing Then
                ResetRollover
            End If
        End If
    Else
    
        If Not m_selectedItem Is Nothing Then
            ResetRollover
        End If
    End If
    
End Function
Public Function SetClearType()
    m_graphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
End Function

Public Function NotifyOptionsChanged()
    PositionOptions
End Function

Private Function PositionOptions()

    Set m_visiblePaneOptions = New Collection

Dim thisOption As Object
Dim thisOptionY As Long

    thisOptionY = TOP_MARGIN

    For Each thisOption In m_navigationPaneOptions

        If thisOption.Visible Then
            thisOption.Position = thisOptionY
            thisOptionY = thisOptionY + m_seperatorAmount
            
            m_visiblePaneOptions.Add thisOption
        End If
    Next

End Function

Public Property Get Image() As GDIPImage
    Set Image = m_surface.Image
End Property

Private Sub pInitialize()

Dim theFont As ViFont
Dim fontSize As Single

    If m_Width = 0 Or m_Height = 0 Then
        MsgBox "Failed to Initialize. Width OR Height is 0. Problems will follow!"
        Exit Sub
    End If
    
    Set m_surface = New GDIPBitmap
    Set m_graphics = New GDIPGraphics
    Set m_Brush = New GDIPBrush
    
    If Layout.GroupMenuSchema.FontID <> "" Then
        Set theFont = Layout.Fonts(Layout.GroupMenuSchema.FontID)
        m_Brush.Colour.Value = theFont.Colour
    Else
        m_Brush.Colour.SetColourByHex "ffffff"
    End If
    
    Set m_font = New GDIPFont
    
    If Not theFont Is Nothing Then
        m_font.Depreciated_Constructor theFont.Face, (theFont.Size * 1.3), FontStyleRegular
    Else
        m_font.Depreciated_Constructor OptionsHelper.PrimaryFont, 12, FontStyleRegular
    End If
    
    
    m_surface.CreateFromSizeFormat m_Width, m_Height, PixelFormat.Format32bppArgb
    m_graphics.FromImage m_surface.Image
    
    m_graphics.CompositingMode = CompositingModeSourceOver
    m_graphics.TextRenderingHint = TextRenderingHintSingleBitPerPixelGridFit
End Sub

Public Property Get Height() As Long
    Height = m_Height
End Property

Private Function Update()

Dim thisOption As Object
Dim thisOptionY As Long

    m_graphics.DrawImageRect m_background, 0, 0, m_Width, m_Height, m_backgroundPosition.X, m_backgroundPosition.Y
    'm_graphics.DrawString "Computer", m_font, m_brush, p

    For Each thisOption In m_visiblePaneOptions
        If thisOption Is m_selectedItem Then
            If m_mouseDown Then
                m_graphics.DrawImageRect m_rollover, m_rolloverPosition.X, m_rolloverPosition.Y, m_rollover.Width, m_rollover.Height / 2, 0, m_rollover.Height / 2
            Else
                m_graphics.DrawImageRect m_rollover, m_rolloverPosition.X, m_rolloverPosition.Y, m_rollover.Width, m_rollover.Height / 2, 0, 0
            End If
            
        End If
        
        If thisOption.OpenAsMenu Then
            'm_graphics.DrawImage m_arrow, (LEFT_MARGIN + m_rollover.Width) - 20, thisOption.Position + ((m_arrow.Height / 2)), m_arrow.Width, m_arrow.Height
            m_graphics.DrawImageRect m_arrow, (LEFT_MARGIN + m_rollover.Width) - 20, thisOption.Position + ((m_arrow.Height / 2)), 4, 7, 0, 0
        End If
        
        m_graphics.DrawString VarScan(thisOption.Caption), m_font, m_Brush, CreatePointF(thisOption.Position, LEFT_MARGIN)
    Next

    
    '---
    
    'theRectF.Width = 100
    'theRectF.Height = 100
    
    'm_Path.Constructor FillModeWinding
    'm_Path.AddString "Computer", m_fontFamily, FontStyleRegular, m_groupMenuFontSize, theRectF, 0
    
    'm_graphics.fillPath m_brush, m_Path
    
End Function

Public Function UpdateBlurred(ByRef blurredBackground As GDIPImage)
    m_graphics.Clear
    m_graphics.DrawImageRect blurredBackground, 0, 0, m_Width, m_Height, m_backgroundPosition.X, m_backgroundPosition.Y
    
    Update
End Function

Public Function UpdateLayered()
    m_graphics.Clear
    
    Update
End Function

Private Function PopulateLinks(ByRef navigationPaneXML As IXMLDOMElement)
    On Error GoTo Handler

Dim thisObjectXML As IXMLDOMElement
Dim thisObject As Object

Dim thisFolderOption As NavigationPaneFolder
Dim thisCustomOption As NavigationPaneCustom

GroupOptionsVisibileCount = 0

MinOSVersion = 0
MaxOSVersion = 0
                
    If navigationPaneXML Is Nothing Then
        Exit Function
    End If
        
    For Each thisObject In navigationPaneXML.childNodes
        
        If TypeName(thisObject) = "IXMLDOMElement" Then
                
            Set thisObjectXML = thisObject

                        If Not IsNull(thisObject.getAttribute("minwinversion")) Then
                                MinOSVersion = Replace(thisObject.getAttribute("minwinversion"), ".", g_WinDecimalSeparator)
                        Else
                                MinOSVersion = Replace(WindowsVersion, ".", g_WinDecimalSeparator)
                        End If
                        
                        If Not IsNull(thisObject.getAttribute("maxwinversion")) Then
                                MaxOSVersion = Replace(thisObject.getAttribute("maxwinversion"), ".", g_WinDecimalSeparator)
                        Else
                                MaxOSVersion = Replace(MinOSVersion, ".", g_WinDecimalSeparator)
                        End If


                        If WindowsVersion >= MinOSVersion And WindowsVersion <= MaxOSVersion Then

                                                        
                                If thisObjectXML.tagName = "folder" Then
                                                Set thisFolderOption = New NavigationPaneFolder
                                                m_navigationPaneOptions.Add thisFolderOption
                                                
                                                thisFolderOption.Visible = True

                                                If Not IsNull(thisObjectXML.getAttribute("visible")) Then thisFolderOption.Visible = CBool(thisObjectXML.getAttribute("visible"))
                                                If thisFolderOption.Visible = True Then GroupOptionsVisibileCount = GroupOptionsVisibileCount + 1
                                                If GroupOptionsVisibileCount > Layout.GroupOptionsVisibilityLimit And Layout.EnableVisibilityLimit Then thisFolderOption.Visible = False
                                                
                                                If Not IsNull(thisObjectXML.getAttribute("display_as_menu")) Then thisFolderOption.OpenAsMenu = CBool(thisObjectXML.getAttribute("display_as_menu"))
                                                If Not IsNull(thisObjectXML.getAttribute("caption")) Then thisFolderOption.Caption = thisObjectXML.getAttribute("caption")
                                                If Not IsNull(thisObjectXML.getAttribute("path")) Then thisFolderOption.Shell = thisObjectXML.getAttribute("path")
                                                If Not IsNull(thisObjectXML.getAttribute("rollover")) Then thisFolderOption.Rollover = thisObjectXML.getAttribute("rollover")

                                ElseIf thisObjectXML.tagName = "custom" Then
                                                Set thisCustomOption = New NavigationPaneCustom
                                                m_navigationPaneOptions.Add thisCustomOption
                                                
                                                thisCustomOption.Visible = True
                                                
                                                If Not IsNull(thisObjectXML.getAttribute("visible")) Then thisCustomOption.Visible = CBool(thisObjectXML.getAttribute("visible"))
                                                If thisCustomOption.Visible = True Then GroupOptionsVisibileCount = GroupOptionsVisibileCount + 1
                                                If GroupOptionsVisibileCount > Layout.GroupOptionsVisibilityLimit And Layout.EnableVisibilityLimit Then thisCustomOption.Visible = False
                                                                                                                
                                                If Not IsNull(thisObjectXML.getAttribute("caption")) Then thisCustomOption.Caption = thisObjectXML.getAttribute("caption")
                                                If Not IsNull(thisObjectXML.getAttribute("rollover")) Then thisCustomOption.Rollover = thisObjectXML.getAttribute("rollover")
                                                If Not IsNull(thisObjectXML.getAttribute("exec")) Then thisCustomOption.Shell = thisObjectXML.getAttribute("exec")
                                                                                                                
                                                thisCustomOption.Populate thisObjectXML
                                End If
                                                        
            End If

            If m_navigationPaneOptions.count = Layout.GroupOptionsLimit Then Exit For
        End If
    Next
        
    Exit Function
Handler:
    Logger.Error Err.Description, "PopulateLinks"
End Function

Private Sub Class_Initialize()
    Set m_logger = LogManager.GetCurrentClassLogger(Me)

    InitializeGDIIfNotInitialized

    Set m_renameTextBox = New VistaSearchBox
    Set m_arrow = New GDIPImage
    
    m_arrow.FromBinary LoadResData("POPOUTMENU_ARROW", "PNG")
    ShowWindow m_renameTextBox.hWnd, SW_HIDE
    
    m_renameTextBox.FocusColour = vbBlack
    m_renameTextBox.BackColour = vbWhite
    m_renameTextBox.Font = g_DefaultFont
End Sub

Private Sub Class_Terminate()
    Set m_renameTextBox = Nothing
End Sub

Private Sub m_contextMenu_onClick(theItemTag As String)
    On Error GoTo Handler
    
    If m_selectedItem Is Nothing Then
        MsgBox "Selected item is null", vbCritical
        Exit Sub
    End If
    
    Select Case UCase$(theItemTag)
    
    Case "OPEN_OPTION"
        RaiseEvent onCommand("explorer " & """" & m_selectedItem.Shell & """")
    
    Case "HIDE_OPTION"
        m_selectedItem.Visible = False
        
        PositionOptions
        RaiseEvent onChanged
        
    Case "SHOW_MENU"
        m_selectedItem.OpenAsMenu = True
        RaiseEvent onChanged
        
    Case "DONT_SHOW_MENU"
        m_selectedItem.OpenAsMenu = False
        RaiseEvent onChanged
        
    Case "RENAME_OPTION"
        RenameSelectedItem
        
    Case "PROPERTIES_OPTION"
        ShellEx VarScan(m_selectedItem.Shell), "properties"
                
        Case "EXECUTE"
                Shell m_selectedItem.Shell, vbNormalFocus
    
        Case "SYSPROPERTIES"
                If WindowsVersion >= 6 Then
                        RaiseEvent onCommand("explorer shell:::{bb06c0e4-d293-4f75-8a90-cb05b6477eee}")
                Else
                        RaiseEvent onCommand(Environ("windir") & "\System32\sysdm.cpl")
                End If

        Case "SEARCH"
                ShowFind
        
        Case "MANAGE"
                RaiseEvent onCommand("compmgmt.msc")

        Case "RUN"
                ShowRun

        Case "SHOWHIDE_D"
                ShowHideDesktopIcons m_selectedItem.Shell
                                        
        Case "SHOWHIDE_C"
                ShowHideComputerIcons m_selectedItem.Shell
        
    End Select
    
    m_contextMenu_onInActive
    Exit Sub
Handler:
    Logger.Error Err.Description, "m_contextMenu_onClick", theItemTag
End Sub


Private Function ShowHideDesktopIcons(ShellItem As String)

    On Error GoTo Handler

        Dim NewStartPanelRegKey As RegistryKey
        Set NewStartPanelRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel")
        If NewStartPanelRegKey Is Nothing Then
            Set NewStartPanelRegKey = Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel")
        End If
        

        Dim DesktopIconStatus As Variant
        DesktopIconStatus = 0
        ' 1 = icon is hidden
        If InStr(1, UCase(ShellItem), g_CLSID_COMPUTER, 1) > 0 Then
                        ' This PC
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_COMPUTER)

                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_COMPUTER, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_COMPUTER, 1)
                        End If

        ElseIf InStr(1, UCase(ShellItem), g_CLSID_USERPROFILE, 1) > 0 Then
                        ' %UserProfile%
                        DesktopIconStatus = m_RegObj.GetValue(g_CLSID_USERPROFILE)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_USERPROFILE, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_USERPROFILE, 1)
                        End If

        ElseIf InStr(1, UCase(ShellItem), g_CLSID_LIBRARIES, 1) > 0 Then
                        ' Libraries
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_LIBRARIES)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_LIBRARIES, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_LIBRARIES, 1)
                        End If

        ElseIf InStr(1, UCase(ShellItem), "450D8FBA-AD25-11D0-98A8-0800361B1103", 1) > 0 Or InStr(1, UCase(ShellItem), "PERSONAL", 1) Or InStr(1, UCase(ShellItem), g_CLSID_MYDOCS, 1) > 0 Then
            ' My Documents - XP
                        ' Always set as Hidden to avoid duplicates
                        DesktopIconStatus = NewStartPanelRegKey.SetValue("{450D8FBA-AD25-11D0-98A8-0800361B1103}", 1)
                        DesktopIconStatus = 0
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYDOCS)

                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYDOCS, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYDOCS, 1)
                        End If

        ElseIf InStr(1, UCase(ShellItem), g_CLSID_NETWORK, 1) > 0 Then
                        ' My network places / Network
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_NETWORK)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_NETWORK, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_NETWORK, 1)
                        End If
                
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_MYPIC, 1) > 0 Or InStr(1, UCase(ShellItem), "MYPICTURES", 1) Then
                        ' Pictures
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYPIC)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYPIC, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYPIC, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_MYMUS, 1) > 0 Or InStr(1, UCase(ShellItem), "MYMUSIC", 1) Then
                        ' Music
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYMUS)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYMUS, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYMUS, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_MYVID, 1) > 0 Or InStr(1, UCase(ShellItem), "MYVIDEO", 1) Then
                        ' Video
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_MYVID)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYVID, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_MYVID, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_DOWNLOADS, 1) > 0 Or InStr(1, UCase(ShellItem), "DOWNLOAD", 1) Then
                        ' Downloads
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_DOWNLOADS)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_DOWNLOADS, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_DOWNLOADS, 1)
                        End If
                        
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_3DOBJECTS, 1) > 0 Then
                        ' 3D Objects
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_3DOBJECTS)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_3DOBJECTS, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_3DOBJECTS, 1)
                        End If
                   
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_GAMES, 1) > 0 Then
                        ' Games Explorer
                        DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_GAMES)
                                        
                        If DesktopIconStatus = 1 Then
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_GAMES, 0)
                        Else
                                DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_GAMES, 1)
                        End If
                                                                        
        ElseIf Left(UCase(ShellItem), 7) = "CONTROL" Then
                        ' Control Panel
                        If WindowsVersion < 6 Then
                                ' XP / 2003
                                If FSO.FolderExists(CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(CreateObject("WScript.Shell").SpecialFolders("Desktop")) & "\" & GetPublicString("strControlPanel") & "." & g_CLSID_CONTROLPANEL) Then
                                        DesktopIconStatus = 1
                                End If
                                
                                If DesktopIconStatus = 1 Then
                                        RmDir CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(CreateObject("WScript.Shell").SpecialFolders("Desktop")) & "\" & GetPublicString("strControlPanel") & "." & g_CLSID_CONTROLPANEL
                                Else
                                        MkDir CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(CreateObject("WScript.Shell").SpecialFolders("Desktop")) & "\" & GetPublicString("strControlPanel") & "." & g_CLSID_CONTROLPANEL
                                End If
                        Else
                ' Vista - 11
                                DesktopIconStatus = NewStartPanelRegKey.GetValue(g_CLSID_CONTROLPANEL)
                                                                
                                If DesktopIconStatus = 1 Then
                                        DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_CONTROLPANEL, 0)
                                Else
                                        DesktopIconStatus = NewStartPanelRegKey.SetValue(g_CLSID_CONTROLPANEL, 1)
                                End If

                        End If
                
        End If
        
        ' Refresh desktop
        Call SHChangeNotify(&H8000000, &H0, 0, 0)
        
        Exit Function
Handler:
    Logger.Error Err.Description, "ShowHideDesktopIcons"
End Function



Private Function ShowHideComputerIcons(ShellItem As String)

On Error GoTo Handler

    Dim HideMyComputerIconsRegKey As RegistryKey
    Set HideMyComputerIconsRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideMyComputerIcons")
        
        Dim ComputerIconStatus As Variant
        ComputerIconStatus = 0
        ' 1 = icon is hidden
                If InStr(1, UCase(ShellItem), g_CLSID_USERPROFILE, 1) > 0 Then
                        ' %UserProfile%
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_USERPROFILE)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_USERPROFILE, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_USERPROFILE, 1)
                        End If

        ElseIf InStr(1, UCase(ShellItem), g_CLSID_LIBRARIES, 1) > 0 Then
                        ' Libraries
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_LIBRARIES)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_LIBRARIES, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_LIBRARIES, 1)
                        End If

        ElseIf InStr(1, UCase(ShellItem), "450D8FBA-AD25-11D0-98A8-0800361B1103", 1) > 0 Or InStr(1, UCase(ShellItem), "PERSONAL", 1) Or InStr(1, UCase(ShellItem), g_CLSID_MYDOCS, 1) > 0 Then
                        ' My Documents - XP
                        ' Always set as Hidden to avoid duplicates
                        ComputerIconStatus = HideMyComputerIconsRegKey.SetValue("{450D8FBA-AD25-11D0-98A8-0800361B1103}", 1)
                        ComputerIconStatus = 0
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYDOCS)

                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYDOCS, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYDOCS, 1)
                        End If

                ElseIf InStr(1, UCase(ShellItem), g_CLSID_MYPIC, 1) > 0 Or InStr(1, UCase(ShellItem), "MYPICTURES", 1) Then
                        ' Pictures
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYPIC)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYPIC, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYPIC, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_MYMUS, 1) > 0 Or InStr(1, UCase(ShellItem), "MYMUSIC", 1) Then
                        ' Music
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYMUS)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYMUS, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYMUS, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_MYVID, 1) > 0 Or InStr(1, UCase(ShellItem), "MYVIDEO", 1) Then
                        ' Video
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_MYVID)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYVID, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_MYVID, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_DOWNLOADS, 1) > 0 Or InStr(1, UCase(ShellItem), "DOWNLOAD", 1) Then
                        ' Downloads
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_DOWNLOADS)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_DOWNLOADS, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_DOWNLOADS, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_3DOBJECTS, 1) > 0 Then
                        ' 3D Objects
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_3DOBJECTS)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_3DOBJECTS, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_3DOBJECTS, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_DESKTOP, 1) > 0 Then
                        ' Desktop
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_DESKTOP)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_DESKTOP, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_DESKTOP, 1)
                        End If
                ElseIf InStr(1, UCase(ShellItem), g_CLSID_GAMES, 1) > 0 Then
                        ' Games Explorer
                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_GAMES)
                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_GAMES, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_GAMES, 1)
                        End If
                                                        
        ElseIf Left(UCase(ShellItem), 7) = "CONTROL" Then
                        ' Control Panel

                        ComputerIconStatus = HideMyComputerIconsRegKey.GetValue(g_CLSID_CONTROLPANEL)
                                                        
                        If ComputerIconStatus = 1 Then
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_CONTROLPANEL, 0)
                        Else
                                ComputerIconStatus = HideMyComputerIconsRegKey.SetValue(g_CLSID_CONTROLPANEL, 1)
                        End If

        End If
        
        ' Refresh explorer
        'Call SHChangeNotify(&H100, &H0, 0, 0)
                
                
        Exit Function
Handler:
    Logger.Error Err.Description, "ShowHideComputerIcons"
End Function

Private Function RenameSelectedItem()

Dim textBoxY As Long
Dim textBoxX As Long

    

    m_windowPosition.Y = m_window.Top / Screen.TwipsPerPixelY
    m_windowPosition.X = m_window.Left / Screen.TwipsPerPixelX
    
    textBoxY = m_windowPosition.Y + m_backgroundPosition.Y + m_rolloverPosition.Y + (m_rolloverHeight / 2) - g_DefaultFont.FontWeight / 2
    textBoxX = m_windowPosition.X + m_backgroundPosition.X + m_rolloverPosition.X + 8
    
    m_renameTextBox.text = m_selectedItem.Caption
    Set m_itemToBeRenamed = m_selectedItem
    
    Call SetWindowPos(m_renameTextBox.hWnd, HWND_TOPMOST, textBoxX, textBoxY, m_rollover.Width - 10, g_DefaultFont.FontWeight, 0)
    
    ShowWindow m_renameTextBox.hWnd, SW_SHOW
    
    SetFocus m_renameTextBox.real_hWnd
    PostMessage ByVal m_renameTextBox.real_hWnd, ByVal EM_SETSEL, ByVal 0, ByVal -1
    
End Function

Private Sub m_customContextMenu_onClick(theItemTag As String)
    RaiseEvent onCommand(theItemTag)
End Sub

Private Sub m_popoutMenu_onInActive()
    m_popoutMenu.Hide
End Sub

Private Sub m_contextMenu_onInActive()
    If Not m_contextMenu Is Nothing Then
        Unload m_contextMenu
        Set m_contextMenu = Nothing
    End If
    
    UpdateMouseAction
End Sub

Private Sub UpdateMouseAction()

    m_windowPosition.Y = m_window.Top / Screen.TwipsPerPixelY
    m_windowPosition.X = m_window.Left / Screen.TwipsPerPixelX
    
Dim pCurrent As POINTL

    GetCursorPos pCurrent
    
    If pCurrent.X > m_windowPosition.X + m_backgroundPosition.X And pCurrent.X < m_windowPosition.X + m_backgroundPosition.X + m_Width And _
        pCurrent.Y > m_windowPosition.Y + m_backgroundPosition.Y And pCurrent.Y < m_windowPosition.Y + m_backgroundPosition.Y + m_Height Then
        
        MouseMove pCurrent
    Else
        MouseOut
    End If
End Sub

Private Function IsContextMenusOpen() As Boolean
    If Not m_contextMenu Is Nothing Then
        IsContextMenusOpen = True
        Exit Function
    End If
    
    If Not m_customContextMenu Is Nothing Then
        If m_customContextMenu.Visible Then
            IsContextMenusOpen = True
            Exit Function
        End If
    End If
End Function

Private Sub m_customContextMenu_onInActive()
    m_customContextMenu.Hide
    
    UpdateMouseAction
End Sub

Private Sub m_popoutMenu_onClickItem(strPath As String)
    SelectBestExecutionMethod strPath
End Sub


Private Sub m_renameTextBox_onKeyDown(KeyCode As Long)
    If KeyCode = vbKeyReturn Then
        m_renameTextBox_onLostFocus
    End If
End Sub

Private Sub m_renameTextBox_onLostFocus()
    If m_itemToBeRenamed Is Nothing Then Exit Sub

    ShowWindow m_renameTextBox.hWnd, SW_HIDE
    m_itemToBeRenamed.Caption = m_renameTextBox.text
    
    RaiseEvent onChanged
End Sub