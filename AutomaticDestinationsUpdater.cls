VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutomaticDestinationsUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lastKnownSizes As Collection
Private m_jumpLists As Collection

Private m_filesFolder As Scripting.Folder

Public Function JumplistsAvailable() As Boolean
    JumplistsAvailable = Not m_filesFolder Is Nothing
End Function

Public Function GetImagePathList(ByVal srcImagePath As String)

Dim thisCollection As Collection
Dim thisPath
Dim theResults() As String
Dim resultIndex As Long
Dim thisImagePath

    srcImagePath = UCase$(StrEnd(srcImagePath, "\"))

    For Each thisCollection In m_jumpLists
    
        For Each thisPath In thisCollection
            thisImagePath = DestermineHandler(CStr(thisPath))
            
            If thisImagePath = srcImagePath Then
                ReDim Preserve theResults(resultIndex)
                theResults(resultIndex) = thisPath
                
                resultIndex = resultIndex + 1
            End If
        Next
    Next
    
    GetImagePathList = theResults

End Function

Private Function UpdateJumplistPaths(theFile As Scripting.File)
Dim newJumplist As Collection

    If ExistInCol(m_jumpLists, theFile.Path) Then
        m_jumpLists.Remove theFile.Path
    End If
    
    Set newJumplist = ParseJumplistFile(theFile.Path)
    m_jumpLists.Add newJumplist, theFile.Path
End Function

Public Function Update()

    On Error GoTo Handler

Dim thisFile As Scripting.File

    For Each thisFile In m_filesFolder.Files
        If Not ExistInCol(m_lastKnownSizes, thisFile.Path) Then
            UpdateJumplistPaths thisFile
            m_lastKnownSizes.Add thisFile.Size, thisFile.Path
        ElseIf CLng(m_lastKnownSizes(thisFile.Path)) <> thisFile.Size Then
            UpdateJumplistPaths thisFile
            
            m_lastKnownSizes.Remove thisFile.Path
            m_lastKnownSizes.Add thisFile.Size, thisFile.Path
        End If
    Next

    Exit Function
Handler:
    Logger.Error Err.Description, "Update"
End Function

Private Sub Class_Initialize()
    
Dim filesPath As String

    Set m_lastKnownSizes = New Collection
    Set m_jumpLists = New Collection
    
    filesPath = Environ$("appdata") & "\Microsoft\Windows\Recent\automaticdestinations\"
    If FSO.FolderExists(filesPath) Then
        Set m_filesFolder = FSO.GetFolder(filesPath)
        Update
    End If
    
End Sub
