VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgramDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_frequentPrograms As Collection
Private m_pinnedPrograms As Collection

Private m_sourceDoc As DOMDocument

Public Event onMetroShortcutAdded()
Public Event onRequestRedraw()

Public Function AddMetroAppsShortcut_ToPinned()
    If ExistsInPinned("explorer shell:::{2559a1f8-21d7-11d4-bdaf-00c04f60b9f0}") Then
        Exit Function
    End If

Dim Program As clsProgram

    Set Program = New clsProgram
    Program.Caption = "Show Metro Apps"
    Program.Path = "explorer shell:::{2559a1f8-21d7-11d4-bdaf-00c04f60b9f0}"
    
    Program.szIcon = AppPath & "plugins\metroprovider.exe"
    Set Program.Icon = IconManager.GetViIcon(Program.szIcon, True)

    'MsgBox "Collection count: " & m_frequentPrograms.Count
    m_pinnedPrograms.Add Program, Program.Path
    
    RaiseEvent onMetroShortcutAdded
End Function

Public Function AddMetroShortcut_ToPinned()
    If ExistsInPinned("!default_menu") Then
        Exit Function
    End If

Dim Program As clsProgram

    Set Program = New clsProgram
    
    If g_Windows8 Or g_Windows81 Then
        Program.Caption = "Show Metro"
    Else
        Program.Caption = "Show Windows Start Menu"
    End If
        
    Program.Path = "!default_menu"
    Set Program.Icon = IconManager.GetViIcon(AppPath & "plugins\metroprovider.exe", True)

    'MsgBox "Collection count: " & m_frequentPrograms.Count
    m_pinnedPrograms.Add Program, Program.Path
    
    RaiseEvent onMetroShortcutAdded
End Function

Public Function GetProgramByPath(szProgramPath As String) As clsProgram

Dim Program As clsProgram

    If ExistCol(m_frequentPrograms, szProgramPath) Then
        Set GetProgramByPath = m_frequentPrograms(szProgramPath)
    ElseIf ExistCol(m_pinnedPrograms, szProgramPath) Then
        Set GetProgramByPath = m_pinnedPrograms(szProgramPath)
    Else
    
        Set Program = New clsProgram
        Set GetProgramByPath = Program
        
        m_frequentPrograms.Add Program, szProgramPath
        
        Program.Path = szProgramPath
        Program.Caption = GetFileName(szProgramPath)
        
        Set Program.Icon = IconManager.GetViIcon(Program.Path, True)
    End If
    
End Function

Public Function ExistsInPinned(theProgramPath As String) As Boolean

    ExistsInPinned = ExistCol(m_pinnedPrograms, theProgramPath)

End Function

Public Function TotalProgramCount()
    TotalProgramCount = m_pinnedPrograms.count + m_frequentPrograms.count
End Function

Public Function TogglePin_ElseAddToPin_ByProgram(ByRef theProgram As clsProgram)
    If theProgram.Caption = vbNullString And theProgram.Path = vbNullString Then
        Exit Function
    End If
    
Dim Program As clsProgram

    'If doesn't exist in either and to pin
    If Not ExistCol(m_pinnedPrograms, theProgram.Path) And Not ExistCol(m_frequentPrograms, theProgram.Path) Then
        
        Set Program = theProgram
        'MsgBox "Collection count: " & m_frequentPrograms.Count
        m_pinnedPrograms.Add Program, Program.Path
        
        Settings.Comit
        
        RaiseEvent onRequestRedraw
    Else
        TogglePin theProgram.Path
    End If

End Function

Public Function RemoveItem(theProgramPath As String)

    If ExistCol(m_pinnedPrograms, theProgramPath) Then
        
        m_pinnedPrograms.Remove theProgramPath
    ElseIf ExistCol(m_frequentPrograms, theProgramPath) Then
        
        m_frequentPrograms.Remove theProgramPath
    End If

    Settings.Comit
End Function

Public Function TogglePin(theProgramPath As String)

Dim theProgram As clsProgram

    If ExistCol(m_pinnedPrograms, theProgramPath) Then
        
        'MsgBox "Exists in Pinned!"
        
        Set theProgram = m_pinnedPrograms(theProgramPath)
        m_pinnedPrograms.Remove theProgram.Path
        
        If theProgram.OpenCount > 0 Then
            m_frequentPrograms.Add theProgram, theProgram.Path
            
            Sort m_frequentPrograms
        End If
        'Sort m_pinnedPrograms

    ElseIf ExistCol(m_frequentPrograms, theProgramPath) Then
        'MsgBox "Already Exist!"
        Set theProgram = m_frequentPrograms(theProgramPath)
    
        m_frequentPrograms.Remove theProgram.Path
        m_pinnedPrograms.Add theProgram, theProgram.Path
        
        Sort m_frequentPrograms
        'Sort m_pinnedPrograms
    End If

    RaiseEvent onRequestRedraw

    Settings.Comit
End Function

Private Function Sort(theCollection As Collection)

Dim EventLoop As Long
Dim intCounter As Long

Dim TempProgram As clsProgram

    For EventLoop = 1 To theCollection.count
        For intCounter = 1 To theCollection.count - 1
            If theCollection(intCounter).OpenCount < theCollection(intCounter + 1).OpenCount Then
                'Swap the memory allocation
                
                Set TempProgram = theCollection(intCounter)
                theCollection.Remove intCounter
                
                theCollection.Add TempProgram, TempProgram.Path, , intCounter
            End If
        Next intCounter
    Next EventLoop
End Function

Public Property Get PinnedPrograms() As Collection
    Set PinnedPrograms = m_pinnedPrograms
End Property

Public Property Get FrequentPrograms() As Collection
    Set FrequentPrograms = m_frequentPrograms
End Property

Public Function UpdateByNode(sourceNode As INode) As Boolean

Dim thisProgram As clsProgram

    Set thisProgram = GetProgramByPath(sourceNode.Tag)
    
    thisProgram.Caption = sourceNode.Caption
    thisProgram.IncreaseCount
    
    If Not sourceNode.Icon Is Nothing Then
        thisProgram.Icon.IconPath = sourceNode.Icon.IconPath
        thisProgram.szIcon = thisProgram.Icon.IconPath
        
        Set thisProgram.Icon = IconManager.GetViIcon(thisProgram.Icon.IconPath, True)
    End If

End Function

Public Function UpdateByProgramPath(sProgramPath As String, Optional bReSort As Boolean = True) As Boolean
    On Error GoTo Handler
    
    'LogError 0, "Update:: " & sProgramPath & " - " & bReSort, "clsProgramDB"
    
Dim Program As clsProgram

    Set Program = GetProgramByPath(sProgramPath)
    Program.IncreaseCount

    If bReSort Then
        Sort m_frequentPrograms
    End If
    
    UpdateByProgramPath = True
    Exit Function
Handler:
    UpdateByProgramPath = False
    
End Function

Public Function DumpPrograms(ByRef sourceDoc As DOMDocument, ByRef parentElement As IXMLDOMElement) As Boolean
    If m_pinnedPrograms Is Nothing Or m_frequentPrograms Is Nothing Then
        Exit Function
    End If

Dim XML_programs As IXMLDOMElement
Dim thisProgram As clsProgram

Dim newItem As IXMLDOMElement

    'Set m_sourceDoc = New DOMDocument
    
    Set XML_programs = sourceDoc.createElement("programs")
    parentElement.appendChild XML_programs
    
    For Each thisProgram In m_pinnedPrograms
        Set newItem = sourceDoc.createElement("pinned_program")
        XML_programs.appendChild newItem
        
        newItem.setAttribute "caption", thisProgram.Caption
        newItem.setAttribute "path", thisProgram.Path
        newItem.setAttribute "opened", thisProgram.OpenCount
        
        If thisProgram.szIcon <> "" Then newItem.setAttribute "icon", thisProgram.szIcon
    Next
    
    For Each thisProgram In m_frequentPrograms
        Set newItem = sourceDoc.createElement("frequent_program")
        XML_programs.appendChild newItem
        
        newItem.setAttribute "caption", thisProgram.Caption
        newItem.setAttribute "path", thisProgram.Path
        newItem.setAttribute "opened", thisProgram.OpenCount
        
        If thisProgram.szIcon <> "" Then newItem.setAttribute "icon", thisProgram.szIcon
    Next

    DumpPrograms = True
    'Set DumpPrograms = XML_programs
    'm_sourceDoc.Save sCon_AppDataPath & "programs.xml"
End Function

Public Function ReadPrograms(ByRef programsXML As IXMLDOMElement)
    On Error GoTo Handler

Dim thisFrequentProgram As IXMLDOMElement
Dim thisProgram As clsProgram
Dim szIconPath As String
Dim MinOSVersion As Double
Dim MaxOSVersion As Double
'Dim exclusiveOSVersion As Single

    Set m_pinnedPrograms = New Collection
    Set m_frequentPrograms = New Collection

    'programsXML.Load sCon_AppDataPath & "programs.xml"
    
    For Each thisFrequentProgram In programsXML.selectNodes("//programs/frequent_program")
        Set thisProgram = New clsProgram
        'Set thisProgram.Icon = New ViIcon
                
        MinOSVersion = 0
        MaxOSVersion = 0
                
        If Not IsNull(thisFrequentProgram.getAttribute("minwinversion")) Then
            MinOSVersion = CSng(thisFrequentProgram.getAttribute("minwinversion"))
                Else
            MinOSVersion = WindowsVersion
        End If
                
        If Not IsNull(thisFrequentProgram.getAttribute("maxwinversion")) Then
            MaxOSVersion = CSng(thisFrequentProgram.getAttribute("maxwinversion"))
                Else
            MaxOSVersion = MinOSVersion
        End If
        
        If WindowsVersion >= MinOSVersion And WindowsVersion <= MaxOSVersion Then
            szIconPath = vbNullString
        
            thisProgram.Path = VarScan(thisFrequentProgram.Attributes.getNamedItem("path").Text)
        
            If FileExists(thisProgram.Path) Then
        
                thisProgram.OpenCount = thisFrequentProgram.Attributes.getNamedItem("opened").Text
                thisProgram.Caption = VarScan(thisFrequentProgram.Attributes.getNamedItem("caption").Text)
               
                If Not IsNull(thisFrequentProgram.getAttribute("icon")) Then
                    thisProgram.szIcon = thisFrequentProgram.getAttribute("icon")
                    szIconPath = VarScan(thisProgram.szIcon)
                End If
            
                If FileExists(szIconPath) Then
                    Set thisProgram.Icon = IconManager.GetViIcon(szIconPath, True)
                Else
                    Set thisProgram.Icon = IconManager.GetViIcon(thisProgram.Path, True)
                    thisProgram.szIcon = vbNullString
                End If
        
                m_frequentPrograms.Add thisProgram, thisProgram.Path
                        
                        End If
                End If
                
    Next

    For Each thisFrequentProgram In programsXML.selectNodes("//programs/pinned_program")
        Set thisProgram = New clsProgram
        
        MinOSVersion = 0
        MaxOSVersion = 0
                
        If Not IsNull(thisFrequentProgram.getAttribute("minwinversion")) Then
            MinOSVersion = CSng(thisFrequentProgram.getAttribute("minwinversion"))
                Else
            MinOSVersion = WindowsVersion
        End If
                
        If Not IsNull(thisFrequentProgram.getAttribute("maxwinversion")) Then
            MaxOSVersion = CSng(thisFrequentProgram.getAttribute("maxwinversion"))
                Else
            MaxOSVersion = MinOSVersion
        End If
        
        If WindowsVersion >= MinOSVersion And WindowsVersion <= MaxOSVersion Then
        
            szIconPath = vbNullString
            
            thisProgram.Path = VarScan(thisFrequentProgram.Attributes.getNamedItem("path").Text)
                        
        If FileExists(thisProgram.Path) Then
                                
            thisProgram.OpenCount = thisFrequentProgram.Attributes.getNamedItem("opened").Text
            thisProgram.Caption = thisFrequentProgram.Attributes.getNamedItem("caption").Text
                                        
                        If Not IsNull(thisFrequentProgram.getAttribute("icon")) Then
                                thisProgram.szIcon = thisFrequentProgram.getAttribute("icon")
                                szIconPath = VarScan(thisProgram.szIcon)
                        End If
                        
                        If FileExists(szIconPath) Then
                                Set thisProgram.Icon = IconManager.GetViIcon(szIconPath, True)
                        Else
                                thisProgram.szIcon = vbNullString
                        End If
                        
                        If thisProgram.Path = "!default_menu" Then
                                Set thisProgram.Icon = IconManager.GetViIcon(AppPath & "plugins\metroprovider.exe", True)
                        Else
                                Set thisProgram.Icon = IconManager.GetViIcon(thisProgram.Path, True)
                        End If

                        m_pinnedPrograms.Add thisProgram, thisProgram.Path
                        
                        End If
                
        End If
    Next
    

    'Sort m_pinnedPrograms
    Sort m_frequentPrograms
    
    Exit Function
Handler:
End Function

Private Sub Class_Initialize()
    Set m_pinnedPrograms = New Collection
    Set m_frequentPrograms = New Collection
End Sub

