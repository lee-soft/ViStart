Attribute VB_Name = "ContextMenuHelper"
Option Explicit

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        m_logger = LogManager.GetLogger("ContextMenuHelper")
    End If
    
    Set Logger = m_logger
End Property

'This module is a generic context menu builder for files shown in ViStart
Public Function BuildGenericFileContextMenu(ByVal szFilePath As String) As frmVistaMenu
    On Error GoTo Handler

Dim newContextMenu As frmVistaMenu

    Set newContextMenu = New frmVistaMenu
    Set BuildGenericFileContextMenu = newContextMenu
    
    newContextMenu.AddItem GetPublicString("strOpen"), "OPEN", True
    newContextMenu.AddItem GetPublicString("strRunAsAdmin"), "RUNASADMIN"
    newContextMenu.AddItem ""
    
    If Settings.Programs.ExistsInPinned(szFilePath) Then
        newContextMenu.AddItem GetPublicString("strUnpinToStartMenu"), "TOGGLEPIN"
    Else
        newContextMenu.AddItem GetPublicString("strPinToStartMenu"), "TOGGLEPIN"
    End If

    If FileExists(szFilePath) Then
        newContextMenu.AddItem ""
        newContextMenu.AddItem GetPublicString("strProperties"), "PROPERTIES"
    End If

    Exit Function
Handler:
    Logger.Error Err.Description, "BuildGenericFileContextMenu"
End Function

Public Function GenericFileContextMenuHandler(ByVal szCommand As String, ByVal szFilePath As String)
    On Error GoTo Handler
        
Dim thisProgram As clsProgram
Dim theBaseMenu As frmStartMenuBase
Dim theRecentPrograms As frmFreq

    Set theBaseMenu = FindFormByName("frmStartMenuBase")
    Set theRecentPrograms = FindFormByName("frmFreq")

    Select Case szCommand
    
    Case "OPEN"
        Settings.Programs.UpdateByProgramPath szFilePath
        SelectBestExecutionMethod szFilePath
        theRecentPrograms.PopulateItems
    
    Case "RUNASADMIN"
        theBaseMenu.CloseMe
        ShellEx szFilePath, "runas"
        
        Settings.Programs.UpdateByProgramPath szFilePath
        theRecentPrograms.PopulateItems
    
    Case "TOGGLEPIN"
        Settings.Programs.TogglePin_ElseAddToPin_ByProgram CreateProgramFromPath(szFilePath)
        theRecentPrograms.PopulateItems
        
    Case "PROPERTIES"
        theBaseMenu.CloseMe
        ShellEx szFilePath, "properties"
        
    End Select
    
    Exit Function
Handler:
    Logger.Error Err.Description, "GenericFileContextMenuHandler"
End Function
