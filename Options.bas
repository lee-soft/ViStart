Attribute VB_Name = "OptionsHelper"
Option Explicit

Public Const POWERMENU_HEIGHT As Long = 233
Public Const POWERMENU_WIDTH As Long = 122

Public bAutoClick As Boolean

Public strIndexingPath As String
Public strNewIndexingPath As String

Public lngIndexingLimit As Long
Public lngNewIndexingLimit As Long

'Properties
Public PrimaryFont As String
Public SecondaryFont As String

Public sFonts() As String

Dim Reg As New clsRegistry

Function StartsWithWindows() As Boolean
    
    StartsWithWindows = False
    
    If LCase$(Registry.Read("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ViStart", "<Empty>")) = LCase$(AppPath & App.EXEName & ".exe") Then
        StartsWithWindows = True
    End If
    
End Function

Public Function ValidateOptions() As Boolean

Dim newSkin As frmSkinSelect

    If Settings.CurrentSkin = vbNullString Then
        Settings.CurrentSkin = "Windows 7 Official Start Menu"
        
        Set newSkin = New frmSkinSelect
        newSkin.Show vbModal
        
        ValidateOptions = ValidateOptions()
        Exit Function
    End If

    If FSO.FolderExists(sCon_AppDataPath & "_skins\" & Settings.CurrentSkin) = False Then
        If FSO.FolderExists(sCon_AppDataPath & "_skins\" & "Windows 7 Official Start Menu") = False Then

            Set newSkin = New frmSkinSelect
            newSkin.Show vbModal
            
            ValidateOptions = ValidateOptions()
            Exit Function
        Else
            If FileCheck(sCon_AppDataPath & "_skins\" & "Windows 7 Official Start Menu") = False Then
                MsgBox "The base skin is broken. Reinstall ViStart", vbCritical
                Exit Function
            End If
            
            Settings.CurrentSkin = "Windows 7 Official Start Menu"
        End If
    End If

    g_resourcesPath = sCon_AppDataPath & "_skins\" & Settings.CurrentSkin & "\"
    
    ValidateOptions = True
End Function

Public Function PutOptions()

    If Not Settings Is Nothing Then Settings.Comit
    If Not MetroUtility Is Nothing Then MetroUtility.DumpOptions
    
End Function

Public Function GetOptions()

Dim strKey() As String
Dim lngKeyIndex As Long

    'bTrayIcon = Registry.GetAppSettingBooleon("Settings\EnableTrayIcon", True)
    bAutoClick = Registry.GetAppSettingBooleon("Settings\EnableAutoClick", True)
    strIndexingPath = Registry.GetAppSetting("Settings\IndexingPath", sVar_Reg_StartMenu_MyDocuments)
    lngIndexingLimit = Registry.GetAppSettingLong("Settings\IndexingLimit", 4096)
End Function

Function GetChildSkins(strPath As String) As Collection

Dim xmlLayout As New DOMDocument
Dim startMenuElement As IXMLDOMElement
Dim childSkins As Collection: Set childSkins = New Collection
Dim child As IXMLDOMElement
Dim startMenuName As String
Dim startMenuID As String
Dim nextStartMenu As CollectionItem

    Set GetChildSkins = childSkins

    If Not FileExists(strPath) Then
        Exit Function
    End If
    
    If Not xmlLayout.Load(strPath) Then
        Exit Function
    End If
    
    For Each child In xmlLayout.selectNodes("startmenus//startmenu_base")
        Set nextStartMenu = New CollectionItem
        nextStartMenu.Value = getAttribute_IgnoreError(child, "name", vbNullString)
        nextStartMenu.Key = getAttribute_IgnoreError(child, "id", vbNullString)

        If nextStartMenu.Value <> vbNullString And nextStartMenu.Key <> vbNullString Then
            childSkins.Add nextStartMenu, nextStartMenu.Key
        End If
    Next
End Function

Private Sub IndexingPath()

    If (strNewIndexingPath <> strIndexingPath) Or _
        (lngNewIndexingLimit <> lngIndexingLimit) Then
        
        strIndexingPath = strNewIndexingPath
        lngIndexingLimit = lngNewIndexingLimit
        
        If g_bIndexing Then
            g_bAbortIndexing = True
        End If
    End If

End Sub
