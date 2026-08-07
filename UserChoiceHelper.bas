Attribute VB_Name = "UserChoiceHelper"
Option Explicit

'Windows 10/11 record per-user default handlers in FileExts\<ext>\UserChoice and
'alternates in OpenWithProgids; the classic HKCR .ext chain is stale for most file
'types there. This maps an app image name (uppercase exe basename, e.g. "NOTEPAD.EXE")
'to the extensions the user's own registry associates with it.

Private Const EXPLORER_FILEEXTS As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"

Private m_extsByImage As Collection

Private m_logger As SeverityLogger

Private Property Get Logger() As SeverityLogger
    If m_logger Is Nothing Then
        Set m_logger = LogManager.GetLogger("UserChoiceHelper")
    End If

    Set Logger = m_logger
End Property

Public Function GetExtensionsForImage(ByVal srcImageName As String) As Collection

    If m_extsByImage Is Nothing Then BuildIndex

    srcImageName = UCase$(srcImageName)
    If ExistInCol(m_extsByImage, srcImageName) Then
        Set GetExtensionsForImage = m_extsByImage(srcImageName)
    Else
        Set GetExtensionsForImage = New Collection
    End If

End Function

Public Function ImageHandlesExtension(ByVal srcImageName As String, ByVal theExtension As String) As Boolean

    If Left$(theExtension, 1) <> "." Then theExtension = "." & theExtension
    ImageHandlesExtension = ExistInCol(GetExtensionsForImage(srcImageName), UCase$(theExtension))

End Function

Private Sub BuildIndex()
    On Error GoTo Handler

Dim fileExts As RegistryKey
Dim extKey As RegistryKey
Dim subKey As RegistryKey

Dim thisNameVar As Variant
Dim progIdVar As Variant
Dim thisExtension As String

    Set m_extsByImage = New Collection

    Set fileExts = Registry.CurrentUser.OpenSubKey(EXPLORER_FILEEXTS)
    If fileExts Is Nothing Then Exit Sub

    For Each thisNameVar In fileExts.GetSubKeyNames
        thisExtension = CStr(thisNameVar)

        If Left$(thisExtension, 1) = "." Then
            Set extKey = fileExts.OpenSubKey(thisExtension)

            If Not extKey Is Nothing Then
                'The user's current default handler
                Set subKey = extKey.OpenSubKey("UserChoice")
                If Not subKey Is Nothing Then
                    AddImage ResolveProgIdToImage(CStr(subKey.GetValue("ProgId", vbNullString))), thisExtension
                End If

                'Registered alternate handlers
                Set subKey = extKey.OpenSubKey("OpenWithProgids")
                If Not subKey Is Nothing Then
                    For Each progIdVar In subKey.GetValueNames
                        AddImage ResolveProgIdToImage(CStr(progIdVar)), thisExtension
                    Next
                End If
            End If
        End If
    Next

    Exit Sub
Handler:
    Logger.Error Err.Description, "BuildIndex"
End Sub

Private Sub AddImage(ByVal theImageName As String, ByVal theExtension As String)

Dim theExtensions As Collection

    If Len(theImageName) = 0 Then Exit Sub

    theExtension = UCase$(theExtension)

    If ExistInCol(m_extsByImage, theImageName) Then
        Set theExtensions = m_extsByImage(theImageName)
    Else
        Set theExtensions = New Collection
        m_extsByImage.Add theExtensions, theImageName
    End If

    If Not ExistInCol(theExtensions, theExtension) Then theExtensions.Add theExtension, theExtension

End Sub

'ProgId -> uppercase exe basename via its open verb. AppX/UWP ProgIds have no
'shell\open\command so they resolve to nothing and are skipped.
Private Function ResolveProgIdToImage(ByVal progId As String) As String
    On Error GoTo Handler

Dim commandKey As RegistryKey
Dim theCommand As String

    If Len(progId) = 0 Then Exit Function

    Set commandKey = Registry.ClassesRoot.OpenSubKey(progId & "\shell\open\command")
    If commandKey Is Nothing Then Exit Function

    theCommand = commandKey.GetValue("", vbNullString)
    If Len(theCommand) = 0 Then Exit Function

    ResolveProgIdToImage = UCase$(StrEnd(Trim$(GetEXEPathFromQuote(GetAbsolutePath(theCommand))), "\"))

    Exit Function
Handler:
    Logger.Error Err.Description, "ResolveProgIdToImage", progId
End Function
