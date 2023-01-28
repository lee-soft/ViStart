Attribute VB_Name = "AutoDesinationHelper"
Option Explicit

Public g_AutomaticDestinationsUpdater As AutomaticDestinationsUpdater

Public Function DestermineHandler(ByVal theFile As String)

Dim theExtension As String

    theExtension = "." & StrEnd(theFile, ".")
    DestermineHandler = UCase$(StrEnd(Trim$(GetEXEPathFromQuote(GetAbsolutePath(GetTypeHandlerPath(theExtension)))), "\"))

End Function

Public Function ParseJumplistFile(sourceJumpListPath As String) As Collection
    On Error Resume Next

Dim file_length As Long

Dim fnum As Long
Dim fileData As String

Dim lastColonSlash As Long
Dim pathStart As Long
Dim pathEnd As Long
Dim theFilePath As String

Dim theResults As Collection
Dim collectionIndex As Long

    Set theResults = New Collection

    Set ParseJumplistFile = theResults
    
    file_length = FileLen(sourceJumpListPath)

    fnum = FreeFile
    fileData = String$(file_length, Chr$(0))

    Open sourceJumpListPath For Binary As #fnum
        Get #fnum, 1, fileData
    Close fnum
    
    'Check for ascii paths
    Do
        lastColonSlash = InStr(lastColonSlash + 1, fileData, ":\")
        If lastColonSlash > 0 Then
            pathStart = InStrRev(fileData, Chr$(0), lastColonSlash) + 1
            pathEnd = InStr(pathStart + 1, fileData, Chr$(0))
            
            theFilePath = Mid$(fileData, pathStart, pathEnd - pathStart)
            If FSO.FileExists(theFilePath) Then
                If Not ExistInCol(theResults, theFilePath) Then theResults.Add theFilePath, theFilePath
            End If
            
        End If

    Loop While lastColonSlash > 0
    
    'Check for unicode paths
    lastColonSlash = 0
    
    Do
        lastColonSlash = InStr(lastColonSlash + 1, fileData, StrConv(":\", vbUnicode))
        If lastColonSlash > 0 Then
            'pathStart = InStrRev(fileData, StrConv(Chr$(0), vbUnicode), lastColonSlash)
            pathStart = FindNextInvalidPathCharacter(fileData, lastColonSlash) + 3
            pathEnd = InStr(pathStart + 1, fileData, StrConv(Chr$(0), vbUnicode))
            
            If pathEnd > 0 Then
                theFilePath = StrConv(Mid$(fileData, pathStart, pathEnd - pathStart), vbFromUnicode)
                
                pathEnd = FindPathEnd(theFilePath) - 1

                If pathEnd > 0 Then
                    theFilePath = Mid$(theFilePath, 1, pathEnd)
                    'Debug.Print pathEnd & ":'" & theFilePath & "'"
                    If Not ExistInCol(theResults, theFilePath) Then theResults.Add theFilePath, theFilePath
                End If
            End If
            
        End If

    Loop While lastColonSlash > 0

End Function

Private Function FindPathEnd(ByVal theSourceString) As Long

Dim lastSlash As Long
Dim partialFilePath As String
Dim charIndex As Long
    
    lastSlash = InStrRev(theSourceString, "\")
    partialFilePath = Mid$(theSourceString, 1, lastSlash)
    charIndex = lastSlash + 1
    
    If lastSlash > 0 Then
        
        While FSO.FileExists(partialFilePath) = False And charIndex < Len(theSourceString)
            partialFilePath = partialFilePath & Mid$(theSourceString, charIndex, 1)
            charIndex = charIndex + 1
        Wend
        
        If FSO.FileExists(partialFilePath) Then
            FindPathEnd = charIndex
        End If
    End If

End Function

Private Function FindNextInvalidPathCharacter(ByRef theSourceString As String, ByVal startPosition As Long) As Long

Dim thisChar As Byte

    Do
        thisChar = Asc(Mid$(theSourceString, startPosition, 1))
        startPosition = startPosition - 1

    Loop While thisChar >= 65 And thisChar <= 90 Or thisChar = 0 Or thisChar = 58

    FindNextInvalidPathCharacter = startPosition

End Function
