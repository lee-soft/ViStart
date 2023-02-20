Attribute VB_Name = "FileVersionInfoHelper"
Option Explicit

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

' API declarations.
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
    dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
    dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
    dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
    dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
    dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
    dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
    dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
    dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
    dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
    dwFileFlagsMask As Long        '  = &h3F for version "0.42"
    dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
    dwFileType As Long             '  e.g. VFT_DRIVER
    dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long           '  e.g. 0
    dwFileDateLS As Long           '  e.g. 0
End Type

Public Function GetVersionInfo(versionFilePath As String) As FileVersionInfo

Dim dummy_handle As Long
Dim lpBuffer As Long

Dim versionInfo_size As Long
Dim versionInfo As VS_FIXEDFILEINFO
Dim Buffer() As Byte
Dim info_size As Long

Dim infoReturn As FileVersionInfo: Set infoReturn = New FileVersionInfo
    Set GetVersionInfo = infoReturn

    info_size = GetFileVersionInfoSize(versionFilePath, dummy_handle)
    If info_size = 0 Then
        Exit Function
    End If

    ReDim Buffer(1 To info_size)

    If Not (GetFileVersionInfo(versionFilePath, 0&, info_size, Buffer(1))) = 0 Then
        If Not VerQueryValue(Buffer(1), "\", lpBuffer, versionInfo_size) = 0 Then
            MoveMemory versionInfo, lpBuffer, Len(versionInfo)

            With infoReturn
                .ProductMajorPart = Format$(versionInfo.dwProductVersionMSh)
                .ProductMinorPart = Format$(versionInfo.dwProductVersionMSl)
                .ProductBuildPart = Format$(versionInfo.dwProductVersionLSh)
            End With
            Exit Function
        End If
    End If
End Function


