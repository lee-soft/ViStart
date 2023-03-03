Attribute VB_Name = "IndexHelper"
Option Explicit

Public g_colSearch As Collection

Private folMyDocuments As Folder
Private g_Dummy As New INode

Public g_bAbortIndexing As Boolean
Public g_bIndexing As Boolean

Private m_old_strItemPath As String

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        m_logger = LogManager.GetLogger("ArrayHelper")
    End If
    
    Set Logger = m_logger
End Property

Private Function GetFolder(strPath As String, ByRef srcFolder As Folder) As Boolean

    On Error GoTo NotFolder:
    Set srcFolder = FSO.GetFolder(strPath)
    
    GetFolder = True
    Exit Function
    
NotFolder:
    GetFolder = False

End Function

Sub Index_MyDirectory()

    On Error GoTo Abort

    CleanCollection g_colSearch
    CleanCollection g_Dummy.Children
    
    g_bAbortIndexing = False

    Set folMyDocuments = FSO.GetFolder(OptionsHelper.strIndexingPath)
    
    If OptionsHelper.lngIndexingLimit > 0 Then
        g_bIndexing = True
        
        IndexFolder folMyDocuments, g_Dummy
        g_bIndexing = False
    End If
    
    If g_bAbortIndexing Then
        g_bAbortIndexing = False
        Index_MyDirectory
    End If

    Exit Sub
Abort:
    If Err.Number = 76 Then
        OptionsHelper.strIndexingPath = sVar_Reg_StartMenu_MyDocuments
        'frmOptions.Show
        
        Index_MyDirectory
    Else
        Logger.Error Err.Description, "Index_MyDirectory" 
    End If
End Sub

Sub CleanCollection(ByRef srcCollection As Collection)

    While srcCollection.count > 0
        srcCollection.Remove 1
    Wend

End Sub

Sub UnIndexFolder(ByRef srcNode As INode)

Dim nodFile As INode

    On Error GoTo Decrepency
    
    For Each nodFile In srcNode.Children
        If nodFile.Children.count > 0 Then
            UnIndexFolder nodFile
        Else
            g_colSearch.Remove nodFile.Tag
        End If
    Next
    
    If ExistInCol(g_colSearch, srcNode.Tag) Then
        g_colSearch.Remove srcNode.Tag
    End If
    
    Exit Sub
    
Decrepency:
    Logger.Error Err.Description. "UnIndexFolder"
End Sub

Sub IndexFolder(ByRef folSource As Folder, ByRef srcNode As INode)

Dim folThis As Folder
Dim fileThis As File

Dim nodFile As INode

    If g_colSearch.count > OptionsHelper.lngIndexingLimit Then
        Exit Sub
    End If

    If g_bAbortIndexing Then
        Exit Sub
    End If

    On Error GoTo Decrepency

    For Each folThis In folSource.SubFolders
        If Not folThis.Attributes And Hidden Then
            Set nodFile = New INode
            
            nodFile.Caption = folThis.Name
            
            nodFile.Tag = "D:" & folThis.Path
            nodFile.Width = 1000
            
            g_colSearch.Add nodFile, nodFile.Tag
            srcNode.Children.Add nodFile, nodFile.Tag
        
            IndexFolder folThis, nodFile
        End If
    Next

    For Each fileThis In folSource.Files
        Set nodFile = New INode
        
        nodFile.Caption = fileThis.Name
        nodFile.SearchIdentifier = MakeSearchable(fileThis.Name)
        
        nodFile.Tag = fileThis.Path
        
        nodFile.Width = 1000
        
        'Mark file as unloaded icon
        nodFile.IconPosition = -3
        nodFile.IsFile = True
        
        g_colSearch.Add nodFile, fileThis.Path
        srcNode.Children.Add nodFile, fileThis.Path
        
        DoEvents
    Next
    
    Exit Sub
    
Decrepency:
    Logger.Error Err.Description, "IndexFolder"
End Sub
