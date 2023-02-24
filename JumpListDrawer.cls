VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JumpListDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public HasMouse As Boolean

Private m_surface As GDIPBitmap
Private m_graphics As GDIPGraphics
Private WithEvents m_contextMenu As frmVistaMenu
Attribute m_contextMenu.VB_VarHelpID = -1

Private m_viFont As ViFont
Private m_font As GDIPFont
Private m_Brush As GDIPBrush

Private m_Width As Long
Private m_Height As Long

Private m_theList As Collection
Private m_Source As JumpList

Private m_rollover As GDIPImage
Private m_selectedItem As JumpListItem

Private m_background As GDIPImage
Private m_backgroundPosition As POINTL

Private m_lastKnownCursorPosition As POINTL

Private Const ITEM_SEPERATOR As Long = 22

Public Event onChanged(ByRef newItem As JumpListItem)
Public Event onMouseExit()

Public Event onRequestClose()

Function ShowContextMenu() As Boolean
    If (m_selectedItem Is Nothing) Then
        ShowContextMenu = False
        Exit Function
    End If
    
    If Not m_contextMenu Is Nothing Then Unload m_contextMenu
    Set m_contextMenu = BuildGenericFileContextMenu(m_selectedItem.Path)
    
    Debug.Print "Attemping Resurrection!"
    m_contextMenu.Resurrect True, FindFormByName("frmStartMenuBase")
End Function

Public Function MouseRightClick(ByVal X As Long, Y As Long)
    ShowContextMenu
End Function

Public Function MouseLeftClick(ByVal X As Long, Y As Long)
    If Not m_selectedItem Is Nothing Then
        If ShellEx(m_selectedItem.Path) = APITRUE Then
            RaiseEvent onRequestClose
        End If
    End If
End Function

Public Function MouseLeaves()
    Debug.Print "MouseLeaves()!"

    If Not m_contextMenu Is Nothing Then
        Exit Function
    End If
    
    Set m_selectedItem = Nothing
    RaiseEvent onMouseExit
End Function

Public Function MouseMove(Position As POINTL)

    m_lastKnownCursorPosition.X = Position.X
    m_lastKnownCursorPosition.Y = Position.Y
    
    If m_theList Is Nothing Then Exit Function
    If Not m_contextMenu Is Nothing Then Exit Function

Dim suggestedIndex As Long
    
    'suggestedIndex = Floor(Position.Y / ITEM_SEPERATOR) + 1
    suggestedIndex = FindIndex(Position.Y, ITEM_SEPERATOR) + 1
    
    If suggestedIndex > 0 And suggestedIndex <= m_theList.count Then
        If Not m_theList(suggestedIndex) Is m_selectedItem Then
            Set m_selectedItem = m_theList(suggestedIndex)
            RaiseEvent onChanged(m_selectedItem)
        End If
    Else
        If Not m_selectedItem Is Nothing Then
            Set m_selectedItem = Nothing
            RaiseEvent onMouseExit 'Achieves the desired effect
        End If
    End If
    
End Function

Public Property Let BackgroundPosition(ByRef newPosition As POINTL)
    m_backgroundPosition = newPosition
End Property

Public Property Let BackgroundSource(ByRef theBackground As GDIPImage)
    Set m_background = theBackground
End Property

Public Property Get Source() As JumpList
    Set Source = m_Source
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Font(newFont As ViFont)
    Set m_viFont = newFont
End Property

Public Property Get Image() As GDIPImage
    Set Image = m_surface.Image
End Property

Public Property Let Size(newSize As SIZEL)
    m_Width = newSize.cx
    m_Height = newSize.cy
    
    pInitialize
End Property

Public Property Let Source(newSource As JumpList)

Dim arrayIndex As Long
Dim theList
Dim thisItem As JumpListItem
Dim currentY As Long

    Set m_Source = newSource
    Set m_theList = New Collection
    
    If g_AutomaticDestinationsUpdater.JumplistsAvailable Then
        'theList = g_AutomaticDestinationsUpdater.GetImagePathList(newSource.ImageName)
        
    End If
    
    theList = ConcatArray(newSource.GetMRUList, theList)
    theList = UniqueValues(theList)
    
    QuickSort_FileAccessed theList
    ReverseArray theList
    
    currentY = ITEM_SEPERATOR

    If IsArrayInitialized(theList) Then
        For arrayIndex = LBound(theList) To UBound(theList)
            If FileExists(CStr(theList(arrayIndex))) Then
                Set thisItem = New JumpListItem
                m_theList.Add thisItem
                
                thisItem.Path = CStr(theList(arrayIndex))
                thisItem.Caption = GetFileName(thisItem.Path)
                Set thisItem.Icon = CreateSmallAlphaIcon(thisItem.Path)
                
                currentY = currentY + ITEM_SEPERATOR
                
                If currentY > m_Height Then
                    Exit For
                End If
            End If
        Next
    End If
End Property

Public Function Update()

Dim textPosition As POINTF
Dim theItem As JumpListItem
Dim yPos As Long
Dim theText As String
Dim widthMinusIcon As String

    If m_theList Is Nothing Then Exit Function
    If m_background Is Nothing Then Exit Function
    
    Debug.Print "Clearing!"
    
    'm_graphics.Clear
    m_graphics.DrawImageRect m_background, 0, 0, m_Width, m_Height, m_backgroundPosition.X, m_backgroundPosition.Y
    
    textPosition.X = 22
    widthMinusIcon = m_Width - textPosition.X
    
    For Each theItem In m_theList
        'Debug.Print theItem
        
        If theItem Is m_selectedItem Then
            m_graphics.DrawImage m_rollover, 0, CSng(yPos), m_rollover.Width, m_rollover.Height
        End If
        
        m_graphics.DrawImage theItem.Icon.Image, 2, CSng(yPos + 2), theItem.Icon.Image.Width, theItem.Icon.Image.Height
        
        textPosition.Y = yPos
        theText = theItem.Caption
        
        If m_graphics.MeasureString(theText, m_font).Width > widthMinusIcon Then
            theText = Mid$(theText, 1, Len(theText) - 3)
            
            While m_graphics.MeasureString(theText & "...", m_font).Width > widthMinusIcon And Len(theText) > 0
                theText = Mid$(theText, 1, Len(theText) - 1)
            Wend
            
            Debug.Print theText
            
            theText = theText & "..."
        End If
        
        m_graphics.DrawString theText, m_font, m_Brush, textPosition
        
        yPos = yPos + ITEM_SEPERATOR
    Next

    Debug.Print "Done!"

End Function

Private Sub pInitialize()
    If m_Width = 0 Or m_Height = 0 Then
        Exit Sub
    End If
    
    Set m_surface = New GDIPBitmap
    Set m_graphics = New GDIPGraphics
    Set m_rollover = New GDIPImage
    
    m_rollover.FromFile ResourcesPath & "jumplist_rollover.png"

    If m_font Is Nothing Then
        Set m_font = New GDIPFont
        
        If Not m_viFont Is Nothing Then
            m_font.Depreciated_Constructor m_viFont.Face, m_viFont.Size, FontStyleRegular
        Else
            m_font.Depreciated_Constructor OptionsHelper.PrimaryFont, 13, FontStyleRegular
        End If
    End If
    Set m_Brush = New GDIPBrush
    
    If Not m_viFont Is Nothing Then
        m_Brush.Colour.Value = m_viFont.Colour
    Else
        m_Brush.Colour.SetColourByHex "2e2e2e"
    End If
    
    m_surface.CreateFromSizeFormat m_Width, m_Height, PixelFormat.Format32bppArgb
    m_graphics.FromImage m_surface.Image
    
    'm_graphics.SmoothingMode = SmoothingModeHighQuality
    m_graphics.CompositingMode = CompositingModeSourceOver
    'm_graphics.CompositingQuality = CompositingQualityHighQuality
    'm_graphics.InterpolationMode = InterpolationModeHighQualityBicubic
    m_graphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
End Sub

Private Sub m_contextMenu_onClick(theItemTag As String)
    Debug.Print "m_contextMenu_onClick!"
    
    m_contextMenu_onInActive
    
    If m_selectedItem Is Nothing Then Exit Sub
    GenericFileContextMenuHandler theItemTag, m_selectedItem.Path
End Sub

Private Sub m_contextMenu_onInActive()
    If Not m_contextMenu Is Nothing Then
        Unload m_contextMenu
        Set m_contextMenu = Nothing
    End If
    
    Me.MouseMove m_lastKnownCursorPosition
    
    Debug.Print "JumpListDrawer::onInActive!"
End Sub
