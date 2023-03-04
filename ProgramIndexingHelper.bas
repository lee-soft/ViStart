Attribute VB_Name = "ProgramIndexingHelper"
'This module populates the treeview containing the program list

Option Explicit

Private xIconDB As DOMDocument30

Private cMerges As New Collection
Private m_exclusionList As Collection

Private FSO As New FileSystemObject

Private iItemCount As Long

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        m_logger = LogManager.GetLogger("ProgramIndexingHelper")
    End If
    
    Set Logger = m_logger
End Property

Public Sub ResetCount()

Dim counterIndex As Long

    For counterIndex = LBound(iLastFileCount) To UBound(iLastFileCount)
        iLastFileCount(counterIndex) = 0
        iLastFolderCount(counterIndex) = 0
    Next


End Sub

Public Sub Initialize()

    Set cMerges = New Collection
    
    Set fFolder_Monitor(0) = FSO.GetFolder(sVar_Reg_StartMenu_CurrentUserPrograms)
    Set fFolder_Monitor(1) = FSO.GetFolder(sVar_Reg_StartMenu_CurrentUser)
    Set fFolder_Monitor(2) = FSO.GetFolder(sVar_Reg_StartMenu_CommonPrograms)
    Set fFolder_Monitor(3) = FSO.GetFolder(sVar_Reg_StartMenu_CommonUser)

    LoadIconDB
End Sub

Public Function GetProgramList() As Collection
    Set GetProgramList = m_exclusionList
End Function

Private Function LoadIconDB() As Boolean

    LoadIconDB = True
    Set xIconDB = New MSXML2.DOMDocument
    
    If xIconDB.Load(StartBarSupport.sCon_AppDataPath & "IconDb_" & CStr(g_sVar_Layout_BackColour) & ".xml") = False Then
        
        LoadIconDB = False
        Exit Function
    End If
    xIconDB.setProperty "SelectionLanguage", "XPath"
    
End Function

Public Sub PopulateStarMenuNodes(ByRef sourceNode As INode)
    
Dim startMenuCommonPrograms As Scripting.Folder
Dim startMenuUserPrograms As Scripting.Folder
Dim windowsDirectory As Scripting.Folder
Dim programFiles As Scripting.Folder

Dim allProgramsCollection As New Collection
    
    Set m_exclusionList = New Collection
    Set windowsDirectory = FSO.GetFolder(Environ$("windir"))

    Set startMenuCommonPrograms = FSO.GetFolder(sVar_Reg_StartMenu_CommonPrograms)
    Set startMenuUserPrograms = FSO.GetFolder(sVar_Reg_StartMenu_CurrentUserPrograms)

    PopulateCollection allProgramsCollection, startMenuCommonPrograms, sVar_Reg_StartMenu_CommonPrograms
    PopulateCollection allProgramsCollection, startMenuUserPrograms, sVar_Reg_StartMenu_CurrentUserPrograms
    
    PopulateCollectionFromEXEDirectory allProgramsCollection, windowsDirectory, "", False

    Set allProgramsCollection = SortFolders(allProgramsCollection)

    PopulateNodeFromCollection sourceNode, allProgramsCollection
    'PopulateNode sourceNode, startMenuCommonPrograms, sVar_Reg_StartMenu_CommonPrograms
    'PopulateNode sourceNode, startMenuUserPrograms, sVar_Reg_StartMenu_CurrentUserPrograms
End Sub

Public Function PopulateCollectionFromEXEDirectory(ByRef Col As Collection, ByRef sourceFolder As Scripting.Folder, Optional ByVal stripperKey As String, Optional includeSubFolders As Boolean = False)

On Error Resume Next

Dim thisFile As Scripting.File
Dim thisFolder As Scripting.Folder
Dim theKey As String
Dim thisColItem As ColItem

    Err.Clear

    For Each thisFile In sourceFolder.Files
        If Err Then
            If Err.Number <> 70 Then
                Logger.Error Err.Description, "PopulateCollectionFromEXEDirectory"
            End If
            
            Err.Clear
        Else
            If Not thisFile.Attributes And Hidden Then
                If LCase$(Right$(thisFile.Name, 3)) = "exe" Then
    
                    theKey = Replace(thisFile.Path, stripperKey, "")
                    If Not ExistInCol(Col, theKey) Then
                        
                    
                        Set thisColItem = New ColItem
                        thisColItem.IsFile = True
                        thisColItem.FileName = thisFile.Name
                        thisColItem.Caption = GetAppDescription(thisFile.Path)

                        thisColItem.Path = thisFile.Path
                        thisColItem.VirtualPath = theKey
                        thisColItem.VisibleOnlyInSearch = True
                        
                        Col.Add thisColItem, theKey
                        m_exclusionList.Add thisColItem, thisColItem.Path
                    End If
                
                End If
            End If
        End If
        
    Next
    
    If Not includeSubFolders Then Exit Function
    
    For Each thisFolder In sourceFolder.SubFolders
        If Not thisFolder.Attributes And Hidden Then
            
        
            theKey = Replace(thisFolder.Path, stripperKey, "")
            If Not ExistInCol(Col, theKey) Then
                Set thisColItem = New ColItem
                thisColItem.IsFile = False
                thisColItem.Caption = thisFolder.Name
                thisColItem.Path = thisFolder.Path
                thisColItem.VirtualPath = theKey
                
                Set thisColItem.Children = New Collection
                
                PopulateCollectionFromEXEDirectory thisColItem.Children, FSO.GetFolder(thisColItem.Path), stripperKey & "\" & thisFolder.Name, includeSubFolders
                Col.Add thisColItem, theKey
            Else

                Set thisColItem = Col(theKey)
                PopulateCollectionFromEXEDirectory thisColItem.Children, FSO.GetFolder(thisFolder.Path), stripperKey & "\" & thisFolder.Name, includeSubFolders
            End If
        End If
    Next

End Function

Public Function PopulateCollection(ByRef Col As Collection, ByRef sourceFolder As Scripting.Folder, Optional ByVal stripperKey As String)
On Error Resume Next

Dim thisFile As Scripting.File
Dim thisFolder As Scripting.Folder
Dim theKey As String
Dim thisColItem As ColItem

    Err.Clear

    For Each thisFile In sourceFolder.Files
        If Err Then
            If Err.Number <> 70 Then
                Logger.Error Err.Description, "PopulateCollection"
            End If
            
            Err.Clear
        Else
            If Not thisFile.Attributes And Hidden Then
    
                theKey = Replace(thisFile.Path, stripperKey, "")
                If Not ExistInCol(Col, theKey) Then
                    Set thisColItem = New ColItem
                    thisColItem.IsFile = True
                    thisColItem.Caption = thisFile.Name
                    thisColItem.Path = thisFile.Path
                    thisColItem.VirtualPath = theKey
                    
                    Col.Add thisColItem, theKey
                    m_exclusionList.Add thisColItem, thisColItem.Path
                End If
            End If
        End If

    Next
    
    For Each thisFolder In sourceFolder.SubFolders
        If Not thisFolder.Attributes And Hidden Then
            
        
            theKey = Replace(thisFolder.Path, stripperKey, "")
            If Not ExistInCol(Col, theKey) Then
                Set thisColItem = New ColItem
                thisColItem.IsFile = False
                thisColItem.Caption = thisFolder.Name
                thisColItem.Path = thisFolder.Path
                thisColItem.VirtualPath = theKey
                
                Set thisColItem.Children = New Collection
                
                PopulateCollection thisColItem.Children, FSO.GetFolder(thisColItem.Path), stripperKey & "\" & thisFolder.Name
                Col.Add thisColItem, theKey
            Else

                Set thisColItem = Col(theKey)
                PopulateCollection thisColItem.Children, FSO.GetFolder(thisFolder.Path), stripperKey & "\" & thisFolder.Name
            End If
        End If
    Next

End Function

Private Function SortFiles(ByRef Col As Collection) As Collection

    Dim colNew As Collection
    
    Dim objCurrent As ColItem
    Dim objCompare As ColItem
    
    Dim lngCompareIndex As Long
    Dim strCurrent As String
    Dim strCompare As String
    Dim blnGreaterValueFound As Boolean

    'make a copy of the collection, ripping through it one item
    'at a time, adding to new collection in right order...
    
    Set colNew = New Collection
    
    For Each objCurrent In Col
        If objCurrent.IsFile Then
            'get value of current item...
            strCurrent = UCase$(objCurrent.Caption)
            
            'setup for compare loop
            blnGreaterValueFound = False
            lngCompareIndex = 0
            
            For Each objCompare In colNew
    
                lngCompareIndex = lngCompareIndex + 1
                strCompare = UCase$(objCompare.Caption)
                
                'we are looking for a string sort...
                If strCurrent < strCompare Then
                    'found an item in compare collection that is greater...
                    'add it to the new collection...
                    blnGreaterValueFound = True
                    colNew.Add objCurrent, , lngCompareIndex
                    Exit For
                End If
            Next
            
            'if we didn't find something bigger, just add it to the end of the new collection...
            If blnGreaterValueFound = False Then
                colNew.Add objCurrent
            End If
        End If
    Next
    
    For Each objCurrent In Col
        If Not objCurrent.IsFile Then
            colNew.Add objCurrent
        End If
    Next

    'return the new collection...
    Set SortFiles = colNew
    Set colNew = Nothing
End Function

Public Function SortFolders(ByRef Col As Collection) As Collection

    Dim colNew As Collection
    
    Dim objCurrent As ColItem
    Dim objCompare As ColItem
    
    Dim lngCompareIndex As Long
    Dim strCurrent As String
    Dim strCompare As String
    Dim blnGreaterValueFound As Boolean

    'make a copy of the collection, ripping through it one item
    'at a time, adding to new collection in right order...
    
    Set colNew = New Collection
    
    For Each objCurrent In Col
        If Not objCurrent.IsFile Then
            'get value of current item...
            strCurrent = UCase$(objCurrent.Caption)
            
            'setup for compare loop
            blnGreaterValueFound = False
            lngCompareIndex = 0
            
            For Each objCompare In colNew
    
                lngCompareIndex = lngCompareIndex + 1
                strCompare = UCase$(objCompare.Caption)
                
                'we are looking for a string sort...
                If strCurrent < strCompare Then
                    'found an item in compare collection that is greater...
                    'add it to the new collection...
                    blnGreaterValueFound = True
                    
                    colNew.Add objCurrent, , lngCompareIndex
                    Exit For
                End If
            Next
            
            'if we didn't find something bigger, just add it to the end of the new collection...
            If blnGreaterValueFound = False Then
                colNew.Add objCurrent
            End If
            
            SortFolders objCurrent.Children
        End If
    Next
    
    For Each objCurrent In Col
        If objCurrent.IsFile Then
            colNew.Add objCurrent
        End If
    Next
    SortFiles Col

    'return the new collection...
    Set SortFolders = colNew
    Set colNew = Nothing
End Function

Public Function GetFileCollection(ByRef sourceFolder As Scripting.Folder, sourceCollection As Collection) As Collection

Dim thisFile As Scripting.File
Dim thisFolder As Scripting.Folder

    For Each thisFile In sourceFolder.Files
        If Not thisFile.Attributes And Hidden Then
            'cRes.createNode ExtOrNot(thisFile.Name), MakeSearchable(thisFile.Name), thisFile.Path, GetIconY(thisFile.Path), Replace(thisFile.Path, stripperKey, "")
            sourceCollection.Add thisFile
        End If
    Next
    
    For Each thisFolder In sourceFolder.SubFolders
        If Not thisFolder.Attributes And Hidden Then
            sourceCollection.Add thisFolder
        End If
    Next

End Function

Public Sub PopulateNodeFromCollection(ByRef sourceNode As INode, ByRef sourceCol As Collection)

Dim cRes As INode
Dim thisColItem As ColItem
Dim newNode As INode

    If sourceCol Is Nothing Then Exit Sub
    
    For Each thisColItem In sourceCol
        If thisColItem.IsFile Then
            Set newNode = sourceNode.createNode(ExtOrNot(thisColItem.Caption), MakeSearchable(thisColItem.Caption), thisColItem.Path, -3, thisColItem.VirtualPath, True)
            
            If LCase$(Right$(GetFileName(thisColItem.Path), 3)) = "lnk" Then
                newNode.EXEName = Replace(UCase$(GetFileName(ResolveLink(thisColItem.Path))), " ", "")
            Else
                newNode.EXEName = Replace(UCase$(GetFileName(thisColItem.Path)), " ", "")
            End If
            
            newNode.visibleInSearchOnly = thisColItem.VisibleOnlyInSearch
        End If
    Next
    
    For Each thisColItem In sourceCol
        If Not thisColItem.IsFile Then
            Set cRes = sourceNode.createNode(thisColItem.Caption, "", thisColItem.Path, -3, thisColItem.VirtualPath, False)
            PopulateNodeFromCollection cRes, thisColItem.Children
        End If
    Next

End Sub

Private Sub PopulateNode(ByRef sourceNode As INode, ByRef sourceFolder As Scripting.Folder, Optional ByVal stripperKey As String)

Dim cRes As INode
Dim thisFile As Scripting.File
Dim thisFolder As Scripting.Folder

    If sourceNode Is Nothing Then Exit Sub
    
    If Replace(sourceFolder.Path, stripperKey, "") = vbNullString Then
        Set cRes = sourceNode
    Else
        Set cRes = sourceNode.createNode(sourceFolder.Name, "", sourceFolder.Path, -3, Replace(sourceFolder.Path, stripperKey, ""), False)
    End If
    
    For Each thisFile In sourceFolder.Files
        If Not thisFile.Attributes And Hidden Then
            cRes.createNode ExtOrNot(thisFile.Name), MakeSearchable(thisFile.Name), thisFile.Path, -3, Replace(thisFile.Path, stripperKey, ""), True
        End If
    Next
    
    For Each thisFolder In sourceFolder.SubFolders
        If Not thisFolder.Attributes And Hidden Then
            PopulateNode cRes, thisFolder, stripperKey
        End If
    Next

End Sub

Public Function FileCountTest()
Dim iFolderCount(0 To 3) As Long
Dim iFileCount(0 To 3) As Long
Dim lngMonitorIndex As Long
Dim bNeedCache As Boolean

    Logger.Trace "Performing a File Count Test", "FileCountTest"

    bNeedCache = False

    For lngMonitorIndex = 0 To 3
        iFolderCount(lngMonitorIndex) = fFolder_Monitor(lngMonitorIndex).SubFolders.count
        iFileCount(lngMonitorIndex) = fFolder_Monitor(lngMonitorIndex).Files.count
    
        If iFileCount(lngMonitorIndex) <> iLastFileCount(lngMonitorIndex) Then
            iLastFileCount(lngMonitorIndex) = iFileCount(lngMonitorIndex)
            
            bNeedCache = True
        End If
        
        If iFolderCount(lngMonitorIndex) <> iLastFolderCount(lngMonitorIndex) Then
            iLastFolderCount(lngMonitorIndex) = iFolderCount(lngMonitorIndex)
            
            bNeedCache = True
        End If
    Next

    FileCountTest = bNeedCache

End Function
