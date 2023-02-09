VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayoutParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public SearchBoxSchema As GenericViElement
Public ProgramMenuSchema As GenericViElement
Public FrequentProgramsMenuSchema As GenericViElement
Public AllProgramsRolloverSchema As GenericViElement
Public AllProgramsArrowSchema As GenericViElement
Public AllProgramsTextSchema As GenericViElement
Public GroupMenuSchema As GenericViElement
Public RolloverPlaceHolder As GenericViElement
Public UserPictureSchema As GenericViElement
Public ShutDownButtonSchema As GenericViElement
Public LogOffButtonSchema As GenericViElement
Public ArrowButtonSchema As GenericViElement
Public FrequentProgramsMenuSeperator As GenericViElement
Public JumpListViewerSchema As GenericViElement
Public ShutDownTextSchema As GenericViElement

Public ViOrb_FullHeight As Boolean

Public ShutDownTextJumpListColour As Long

Public ProgramMenuColour As Long
Public FrequentProgramsMenuColour As Long
Public FrequentProgramsSeperatorColour As Long

Public GroupOptionsSeparator As Long
Public SearchBoxFocusColour As Long
Public SearchBoxForeColour As Long
Public YOffset As Long
Public XOffset As Long

Public ProgramsMenuSeperatorColour As Long
Public GroupOptionsLimit As Long
Public GroupOptionsVisibilityLimit As Long
Public EnableVisibilityLimit As Boolean
Public ForceClearType As Boolean

Public ErrorParsing As Boolean


Public Fonts As Collection

Private Function LoadFromResources(ByRef sourceXml As DOMDocument) As Boolean

Dim defaultLayout As String
    defaultLayout = LoadStringFromResource("default_layout.xml", "SCHEMA")
    
    If sourceXml.loadXML(defaultLayout) = False Then
        CreateError "LayoutParser", "LoadFromResources", "Couldn't parse layout file within executable"
        MsgBox "Couldn't parse layout file within executable", vbCritical, "XML Syntax Error"
        Exit Function
    End If
    
    LoadFromResources = True
End Function

Function ParseStartMenu(strPath As String, Optional strStartMenuID As String = "") As StartMenuParseResult
    On Error Resume Next

Dim result As StartMenuParseResult: Set result = New StartMenuParseResult
Set ParseStartMenu = result

Dim xmlLayout As New DOMDocument

Dim startMenuElement As IXMLDOMElement
Dim subElement As IXMLDOMElement
Dim nullElement As IXMLDOMElement

Dim defaultLayoutBinary() As Byte
Dim defaultLayout As String

Dim thisFont As ViFont
Dim defaultFontFace As String
    
Dim validLayout As Boolean
Dim multiLayout As Boolean

    result.ErrorCode = 300
    
    If FileExists(strPath) Then
        If xmlLayout.Load(strPath) = False Then
            CreateError "LayoutParser", "ParseStartMenu", "Couldn't parse layout.xml file within resources"
            MsgBox "Couldn't parse layout.xml file within resources folder", vbCritical, "XML Syntax Error"
            ParseStartMenu = False
            Exit Function
        End If
    Else
        If Not LoadFromResources(xmlLayout) Then
            ParseStartMenu = False
            Exit Function
        End If
    End If

    ForceClearType = False
    ProgramMenuColour = vbWhite
    FrequentProgramsMenuColour = vbWhite
    SearchBoxForeColour = HEXCOL2RGB("#6D6D6D")
    SearchBoxFocusColour = HEXCOL2RGB("#000000")
    FrequentProgramsClientColour = HEXCOL2RGB("#464646")
    ProgramsMenuSeperatorColour = HEXCOL2RGB("#1e3287")
    FrequentProgramsSeperatorColour = HEXCOL2RGB("#D6E4F5")
    YOffset = -4
    XOffset = 0
    
    GroupOptionsSeparator = 35
    GroupOptionsLimit = 18
    
    defaultFontFace = g_DefaultFont.FontFace
    
    validLayout = True
    
    If Not xmlLayout.selectSingleNode("startmenus") Is Nothing Then
        multiLayout = True
    End If
    
    If multiLayout Then
        If strStartMenuID <> vbNullString Then
            Set startMenuElement = xmlLayout.selectSingleNode("startmenus//startmenu_base[@id='" & strStartMenuID & "']")
        Else
            Set allStartMenuBases = xmlLayout.selectNodes("startmenus//startmenu_base")
            Set startMenuElement = allStartMenuBases.nextNode
        End If
    Else
        Set startMenuElement = xmlLayout.selectSingleNode("startmenu_base")
    End If
    
    If startMenuElement Is Nothing Then
        MsgBox "Unable to select a start menu from layout", vbCritical
        Exit Function
    End If

    For Each subElement In startMenuElement.selectNodes("vifont")

        Set thisFont = New ViFont

        'Defaults
        thisFont.Colour = vbBlack
        
        thisFont.Face = g_DefaultFont.FontFace
        thisFont.Size = 15
        
        thisFont.Colour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "colour", "#ffffff"))
        thisFont.Face = getAttribute_IgnoreError(subElement, "face", defaultFontFace)
        
        If Not FontExists(thisFont.Face) Then
            thisFont.Face = defaultFontFace
        End If
        
        thisFont.Size = getAttribute_IgnoreError(subElement, "size", 15)
                             
        Fonts.Add thisFont, getAttribute_IgnoreError(subElement, "id", "fnt_" & Fonts.count)
    Next
    
    For Each subElement In startMenuElement.selectNodes("vielement")
        
        Select Case subElement.getAttribute("id")
        
        Case "searchbox"
            Set SearchBoxSchema = ViElementFromXML(subElement)
            validLayout = IIf(ValidateElement(subElement.getAttribute("id"), SearchBoxSchema) = False, False, validLayout)

        Case "programmenu"
            Set ProgramMenuSchema = ViElementFromXML(subElement)
            validLayout = IIf(ValidateElement(subElement.getAttribute("id"), ProgramMenuSchema) = False, False, validLayout)

        Case "frequentprogramsmenu"
            Set FrequentProgramsMenuSchema = ViElementFromXML(subElement)
            validLayout = IIf(ValidateElement(subElement.getAttribute("id"), FrequentProgramsMenuSchema) = False, False, validLayout)
            
        Case "allprograms_rollover"
            Set AllProgramsRolloverSchema = ViElementFromXML(subElement)
        Case "allprograms_arrow"
            Set AllProgramsArrowSchema = ViElementFromXML(subElement)
            
        Case "allprograms_text"
            Set AllProgramsTextSchema = ViElementFromXML(subElement)
            validLayout = IIf(ValidateElement(subElement.getAttribute("id"), AllProgramsTextSchema) = False, False, validLayout)
            
        Case "groupoptions"
            Set GroupMenuSchema = ViElementFromXML(subElement)

            If GroupMenuSchema.Width = -1 Then
                GroupMenuSchema.Left = GroupMenuSchema.Left - 3
                GroupMenuSchema.Top = GroupMenuSchema.Top - 3
                
                GroupMenuSchema.Width = 140
                GroupMenuSchema.Height = Screen.Height / Screen.TwipsPerPixelY
            End If
            
            validLayout = IIf(ValidateElement(subElement.getAttribute("id"), GroupMenuSchema) = False, False, validLayout)
            
        Case "rolloverplaceholder"
            Set RolloverPlaceHolder = ViElementFromXML(subElement)
        Case "userpicture"
            Set UserPictureSchema = ViElementFromXML(subElement)
            
        Case "shutdown_button"
            Set ShutDownButtonSchema = ViElementFromXML(subElement)
        Case "logoff_button"
            Set LogOffButtonSchema = ViElementFromXML(subElement)
        Case "arrow_button"
            Set ArrowButtonSchema = ViElementFromXML(subElement)
            
        Case "jumplist_viewer"
            Set JumpListViewerSchema = ViElementFromXML(subElement)
            
        Case "shutdown_text"
            Set ShutDownTextSchema = ViElementFromXML(subElement)
            validLayout = IIf(ValidateElement(subElement.getAttribute("id"), ShutDownTextSchema) = False, False, validLayout)
            
        End Select
    Next
    
    If UserPictureSchema Is Nothing Then
        Set UserPictureSchema = RolloverPlaceHolder
    End If
    
    Set subElement = startMenuElement.selectSingleNode("startmenu_base")
    
    result.StartMenuPath = getAttribute_IgnoreError(subElement, "background", "startmenu.png")
    result.AllProgramsPath = getAttribute_IgnoreError(subElement, "allprograms", "allprograms.png")
    result.BottomButtonsArrowPath = getAttribute_IgnoreError(subElement, "bottombuttons_arrow", "bottombuttons_arrow.png")
    result.ButtonPath = getAttribute_IgnoreError(subElement, "bottombuttons_logoff", "bottombuttons_logoff.png")
    result.ProgramsArrowPath = getAttribute_IgnoreError(subElement, "programs_arrow", "programs_arrow.png")
    result.UserFramePath = getAttribute_IgnoreError(subElement, "userframe", "userframe.png")
    result.BottomButtonsShutdownPath = getAttribute_IgnoreError(subElement, "bottombuttons_shutdown", "bottombuttons_shutdown.png")
    
    result.StartMenuExpandedPath = getAttribute_IgnoreError(subElement, "background_expanded", "startmenu_expanded.png")
    result.StartMenuMaskPath = getAttribute_IgnoreError(subElement, "background_masked", "startmenu_mask.bmp")
    
    If Not IsNull(subElement.getAttribute("force_clear_type")) Then ForceClearType = CBool(subElement.getAttribute("force_clear_type"))
    If Not IsNull(subElement.getAttribute("y_offset")) Then YOffset = CLng(subElement.getAttribute("y_offset"))
    If Not IsNull(subElement.getAttribute("x_offset")) Then XOffset = CLng(subElement.getAttribute("x_offset"))
    
    Set subElement = startMenuElement.selectSingleNode("viorb")
    ViOrb_FullHeight = CBool(getAttribute_IgnoreError(subElement, "fullheight_superbar", False))

    Set subElement = startMenuElement.selectSingleNode("vielement[@id='programmenu']")
    ProgramMenuColour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "backcolour", "#ffffff"))
    
    Set subElement = startMenuElement.selectSingleNode("vielement[@id='frequentprogramsmenu']")
    FrequentProgramsMenuColour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "backcolour", "#ffffff"))
    
    Set subElement = startMenuElement.selectSingleNode("groupoptions")
    If Not IsNull(subElement.getAttribute("separator")) Then GroupOptionsSeparator = CLng(subElement.getAttribute("separator"))
        
        GroupOptionsVisibilityLimit = GroupOptionsLimit
    If Not IsNull(subElement.getAttribute("limit")) Then GroupOptionsVisibilityLimit = CLng(subElement.getAttribute("limit"))
    If Not IsNull(subElement.getAttribute("enabled")) Then EnableVisibilityLimit = CBool(getAttribute_IgnoreError(subElement, "enabled", False))
        
    Set subElement = startMenuElement.selectSingleNode("searchbox")
    SearchBoxFocusColour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "focuscolour", "#000000"))
    SearchBoxForeColour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "forecolour", "#6D6D6D"))
    
    Set subElement = startMenuElement.selectSingleNode("programmenu")
    ProgramsMenuSeperatorColour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "separatorfontcolour", "#1e3287"))

    Set subElement = startMenuElement.selectSingleNode("frequentprogramsmenu")
    If Not IsNull(subElement.getAttribute("separatorcolour")) Then
        FrequentProgramsSeperatorColour = HEXCOL2RGB(subElement.getAttribute("separatorcolour"))
    End If

    Set subElement = startMenuElement.selectSingleNode("shutdowntext")
    
    If getAttribute_IgnoreError(subElement, "jumplistcolour", "empty") <> "empty" Then
        ShutDownTextJumpListColourSet = True
        ShutDownTextJumpListColour = HEXCOL2RGB(subElement.getAttribute("jumplistcolour"))
    Else
        Set subElement = startMenuElement.selectSingleNode("vielement[@id='shutdown_text']")
        ShutDownTextJumpListColour = HEXCOL2RGB(getAttribute_IgnoreError(subElement, "jumplistcolour", "#000000"))
    End If
    
    ParseStartMenu = validLayout
    result.ErrorCode = 0
End Function

Private Function ValidateElement(ByVal theElementId As String, ByRef theElement As GenericViElement) As Boolean

    If theElement.FontID = "" Then
        ValidateElement = True
        Exit Function
    End If

    If Not ExistInCol(Me.Fonts, theElement.FontID) Then
        ErrMessage theElementId, "Couldn't find the font specified"
        Exit Function
    End If

    ValidateElement = True
End Function

Private Function ErrMessage(ByVal theElementId As String, ByVal theDescription As String)
    MsgBox "There was a problem with '" & theElementId & "' element in layout.xml" & vbCrLf & _
            vbCrLf & _
            "Reason: " & theDescription, vbCritical, "Error parsing layout.xml file"
End Function


Private Sub Class_Initialize()
    Set Fonts = New Collection
    'ErrorParsing = ParseStartMenu(ResourcesPath & "layout.xml")
End Sub
