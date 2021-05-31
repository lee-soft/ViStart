VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Programs As clsProgramDB
Public NavigationPane As ViNavigationPane

Private m_settingsDoc As DOMDocument

Public navigationXML As IXMLDOMElement
Public programsXML As IXMLDOMElement
Private m_generalXML As IXMLDOMElement

Public CurrentSkin As String
Public CurrentOrb As String

Public EnableAutoClick As Boolean
Public EnableTrayIcon As Boolean
Public ShowUserPicture As Boolean
Public ShowProgramsFirst As Boolean
Public StartButtonShowsWindowsMenu As Boolean
Public CatchLeftWindowsKey As Boolean
Public CatchRightWindowsKey As Boolean
Public LanguageFile As String

'Public IndexingLimit As Boolean
'Public IndexingPath As String

Private Sub Class_Initialize()
    On Error GoTo Handler
    SetVars_IfNeeded
    
Dim settingsXML As IXMLDOMElement

    Set m_settingsDoc = New DOMDocument
    
    Set NavigationPane = New ViNavigationPane
    Set Programs = New clsProgramDB

    If Not m_settingsDoc.Load(sCon_AppDataPath & "settings.xml") Then
        m_settingsDoc.loadXML LoadStringFromResource("default_settings.xml", "SCHEMA")
    End If
    
    Set settingsXML = m_settingsDoc.firstChild
    
    Set navigationXML = settingsXML.selectSingleNode("//navigation_pane")
    Set programsXML = settingsXML.selectSingleNode("//programs")
    Set m_generalXML = settingsXML.selectSingleNode("//general")
    
    Me.CurrentSkin = CStr(ExtractXMLTextElement(m_generalXML, "current_skin", ""))
    Me.CurrentOrb = CStr(ExtractXMLTextElement(m_generalXML, "current_orb", ""))
    Me.EnableTrayIcon = CBool(ExtractXMLTextElement(m_generalXML, "enable_trayicon", CStr(True)))
    Me.ShowUserPicture = CBool(ExtractXMLTextElement(m_generalXML, "show_userpicture", CStr(True)))
    Me.ShowProgramsFirst = CBool(ExtractXMLTextElement(m_generalXML, "show_program_menu_first", CStr(False)))
    Me.StartButtonShowsWindowsMenu = CBool(ExtractXMLTextElement(m_generalXML, "start_button_shows_windows_menu", CStr(False)))
    Me.CatchLeftWindowsKey = CBool(ExtractXMLTextElement(m_generalXML, "catch_left_windows_key", CStr(True)))
    Me.CatchRightWindowsKey = CBool(ExtractXMLTextElement(m_generalXML, "catch_right_windows_key", CStr(True)))
    Me.LanguageFile = CStr(ExtractXMLTextElement(m_generalXML, "language", vbNullString))
    
    If Not FileExists(App.Path & "\languages\" & Me.LanguageFile & ".xml") Then
        If FileExists(App.Path & "\languages\" & LocaleHelper.GetUserLocale() & ".xml") Then
            Me.LanguageFile = LocaleHelper.GetUserLocale()
        Else
            If FileExists(App.Path & "\languages\english.xml") Then
                LogError "Native Language: " & LocaleHelper.GetUserLocale() & " unavailable, defaulting to American English..", "ViSettings::Initialize()"
                Me.LanguageFile = "English"
            Else
                LogError "No languages installed.. using precompiled values"
            End If
        End If
    End If

    PopulateUserStringsFromXML App.Path & "\languages\" & Me.LanguageFile & ".xml"
    Programs.ReadPrograms programsXML
    
Handler:
End Sub

Public Function Comit()

Dim settingsXML As IXMLDOMElement
Dim generalXML As IXMLDOMElement
'Dim navigationXML As IXMLDOMElement

    Set m_settingsDoc = New DOMDocument
    
    Set settingsXML = m_settingsDoc.createElement("settings")
    Set generalXML = m_settingsDoc.createElement("general")
    
    m_settingsDoc.appendChild settingsXML
    settingsXML.appendChild generalXML

    CreateXMLTextElement m_settingsDoc, generalXML, "current_skin", Me.CurrentSkin
    CreateXMLTextElement m_settingsDoc, generalXML, "current_orb", Me.CurrentOrb
    CreateXMLTextElement m_settingsDoc, generalXML, "enable_trayicon", CStr(Me.EnableTrayIcon)
    CreateXMLTextElement m_settingsDoc, generalXML, "show_userpicture", CStr(Me.ShowUserPicture)
    CreateXMLTextElement m_settingsDoc, generalXML, "show_program_menu_first", CStr(Me.ShowProgramsFirst)
    CreateXMLTextElement m_settingsDoc, generalXML, "start_button_shows_windows_menu", CStr(Me.StartButtonShowsWindowsMenu)
    CreateXMLTextElement m_settingsDoc, generalXML, "catch_left_windows_key", CStr(Me.CatchLeftWindowsKey)
    CreateXMLTextElement m_settingsDoc, generalXML, "catch_right_windows_key", CStr(Me.CatchRightWindowsKey)
    CreateXMLTextElement m_settingsDoc, generalXML, "language", CStr(Me.LanguageFile)

    If Not Programs.DumpPrograms(m_settingsDoc, settingsXML) Then
        Exit Function
    End If
    
    If Not NavigationPane.DumpOptions(m_settingsDoc, settingsXML) Then
        Exit Function
    End If

    m_settingsDoc.Save sCon_AppDataPath & "settings.xml"
End Function

Private Sub Class_Terminate()
    Comit
End Sub
