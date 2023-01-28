Attribute VB_Name = "ThemeHelper"
Option Explicit

'//source was in C# from urls:
'//http://www.codeproject.com/csharp/CompressWithWinShellAPICS.asp
'//http://www.codeproject.com/csharp/DecompressWinShellAPICS.asp

'//set reference to "Microsoft Shell Controls and Automation"


'http://forums.microsoft.com/MSDN/ShowPost.aspx?PostID=1090552&SiteID=1
'Be aware when using the shell automation interface to unzip files as it
'leaves copies of the zip files in the temp directory (defined by %TEMP%).
'Folders named "Temporary Directory X for demo.zip" are generated where X
'is a sequential number from 1 - 99.  When it reaches 99 you will then get
'a error dialog saying "The file exists" and it will not continue.
'I 've no idea why Windows doesn't clean up after itself when unzipping files,
'but it is most annoying...


'//CopyHere options
'0 Default. No options specified.
'4 Do not display a progress dialog box.
'8 Rename the target file if a file exists at the target location with the same name.
'16 Click "Yes to All" in any dialog box displayed.
'64 Preserve undo information, if possible.
'128 Perform the operation only if a wildcard file name (*.*) is specified.
'256 Display a progress dialog box but do not show the file names.
'512 Do not confirm the creation of a new directory if the operation requires one to be created.
'1024 Do not display a user interface if an error occurs.
'4096 Disable recursion.
'9182 Do not copy connected files as a group. Only copy the specified files.

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function InstallOrb(szSourceFile As String, ByRef szOrbName As String) As Boolean

Dim fileNameHolder As String
Dim newFile As Scripting.File

    If FileExists(szSourceFile) Then
        Set newFile = FSO.GetFile(szSourceFile)
 
        'makesure we're not trying to install an Orb that's already installed
        If LCase$(IIf(Right$(newFile.ParentFolder, 1) = "\", newFile.ParentFolder, newFile.ParentFolder & "\")) <> LCase$(sCon_OrbFolderPath) Then
            szOrbName = PathFindFileName(szSourceFile)

            If Not FileExists(sCon_OrbFolderPath & szOrbName) Then
                fileNameHolder = sCon_OrbFolderPath & szOrbName
                If fileNameHolder = vbNullString Then Exit Function
            
                FSO.CopyFile szSourceFile, fileNameHolder, True
            Else
                fileNameHolder = sCon_OrbFolderPath & szOrbName
            End If
                
            Set newFile = FSO.GetFile(fileNameHolder)
        Else
            fileNameHolder = newFile.Path
        End If
        
        szOrbName = PathFindFileName(fileNameHolder)
        InstallOrb = True
    End If

End Function

Public Function InstallTheme(szSourceFile As String, ByRef szThemeName As String) As Boolean
    szThemeName = ExtOrNot(PathFindFileName(szSourceFile))
    Zip_Activity "UNZIP", szSourceFile, sCon_AppDataPath & "_skins\" & szThemeName
    
    InstallTheme = FSO.FileExists(sCon_AppDataPath & "_skins\" & szThemeName & "\startmenu.png")
End Function

Public Function PathIsDirectory(ByVal szPath As String) As Boolean
    PathIsDirectory = FSO.FolderExists(szPath)
End Function

'Removes the trailing backslash from a given path (if it has one)
Public Function PathRemoveBlackSlash(ByVal szPath As String)

    If Right$(szPath, 1) = "\" Then
        PathRemoveBlackSlash = Left$(szPath, Len(szPath) - 1)
    Else
        PathRemoveBlackSlash = szPath
    End If

End Function

Public Function PathFindExtension(ByVal szPath As String)

Dim periodPosition As Long

    periodPosition = InStrRev(szPath, ".")
    If periodPosition > 0 Then
        PathFindExtension = Right$(szPath, Len(szPath) - periodPosition)
    End If

End Function

Public Function PathFindFileName(ByVal szPath As String) As String

Dim lastBackSlashPosition As Long

    lastBackSlashPosition = InStrRev(szPath, "\")

    If lastBackSlashPosition = Len(szPath) Then
        lastBackSlashPosition = InStrRev(Mid$(szPath, 1, Len(szPath) - 1), "\")
    End If

    If lastBackSlashPosition = 0 Then
        PathFindFileName = szPath
        Exit Function
    End If
    
    PathFindFileName = Right$(szPath, Len(szPath) - lastBackSlashPosition)
End Function

Sub Zip_Activity(Action As String, sFileSource As String, sFileDest As String)

Dim thisFile As Scripting.File
Dim originalName As String
Dim oShell As Object
Dim fileSource As Object
Dim fileDest As Object

    On Error GoTo EH

    If FSO.FolderExists(sFileDest) = False Then
        FSO.CreateFolder sFileDest
    End If

    Set thisFile = FSO.GetFile(sFileSource)
    originalName = thisFile.Name
    
    If LCase$(Right$(thisFile.Name, 4)) <> ".zip" Then
        originalName = thisFile.Name
        thisFile.Name = thisFile.Name & ".zip"
    End If

    If sFileSource = "" Or sFileDest = "" Then GoTo EH
    
    Set oShell = CreateObject("Shell.Application")
    If oShell Is Nothing Then GoTo EH
    
    Select Case UCase$(Action)
    
        Case "UNZIP"
            
            If Right$(UCase$(sFileSource), 4) <> ".ZIP" Then
                sFileSource = sFileSource & ".ZIP"
            End If
            
            Set fileSource = oShell.NameSpace("" & sFileSource)      '//should be zip file
            Set fileDest = oShell.NameSpace("" & sFileDest)          '//should be directory

            Call fileDest.CopyHere(fileSource.Items, 20)
        
        Case Else
        
    End Select
    
    If thisFile.Name <> originalName Then thisFile.Name = originalName

            
    '//Ziping a file using the Windows Shell API creates another thread where the zipping is executed.
    '//This means that it is possible that this console app would end before the zipping thread
    '//starts to execute which would cause the zip to never occur and you will end up with just
    '//an empty zip file. So wait a second and give the zipping thread time to get started.

    Call Sleep(1000)

EH:
    Set oShell = Nothing
    Set fileSource = Nothing
    Set fileDest = Nothing
    Exit Sub

    If Err.Number = 70 Then
        MsgBox "ViStart does not have exclusive access to the skin. If you have the skin file open in another program please close it", vbCritical
    Else
        MsgBox "There was a problem installing the skin." & vbCrLf & "Makesure the skin file isn't open by another program!", vbExclamation, "error"
        
        On Error Resume Next
        If originalName <> vbNullString Then
            If thisFile.Name <> originalName Then thisFile.Name = originalName
        End If
    End If
End Sub

Private Function Create_Empty_Zip(sFileName As String) As Boolean

    Dim EmptyZip()  As Byte
    Dim J           As Integer

    On Error GoTo EH
    Create_Empty_Zip = False

    '//create zip header
    ReDim EmptyZip(1 To 22)

    EmptyZip(1) = 80
    EmptyZip(2) = 75
    EmptyZip(3) = 5
    EmptyZip(4) = 6
    
    For J = 5 To UBound(EmptyZip)
        EmptyZip(J) = 0
    Next

    '//create empty zip file with header
    Open sFileName For Binary Access Write As #1

    For J = LBound(EmptyZip) To UBound(EmptyZip)
        Put #1, , EmptyZip(J)
    Next
    
    Close #1

    Create_Empty_Zip = True

EH:
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "Error"
    End If
    
End Function
