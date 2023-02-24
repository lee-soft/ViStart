VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" _
    (ByVal grpnm As Long, ByVal parnm As Long, _
    ByVal deflt As Long, ByVal parvl As Long, _
    ByVal parlen As Long, ByVal INIPath As Long) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" _
    (ByVal grpnm As Long, ByVal parnm As Long, _
    ByVal parvl As Long, ByVal INIPath As Long) As Long

Private m_INIPath As String

Public Property Let INIPath(newIniPath As String)
    
Dim hFile As Long
Dim wBOM(1) As Byte
Dim bytesWritten As Long
    
    If Not FileExists(newIniPath) Then
        wBOM(0) = &HFF
        wBOM(1) = &HFE

        hFile = CreateFile(ByVal newIniPath, ByVal GENERIC_WRITE, ByVal 0, _
                   ByVal 0, ByVal CREATE_NEW, ByVal FILE_ATTRIBUTE_NORMAL, ByVal 0)
                   
        WriteFile ByVal hFile, ByVal VarPtr(wBOM(0)), ByVal 2, bytesWritten, ByVal 0
        CloseHandle hFile
    End If
    
    m_INIPath = newIniPath
    
End Property

Public Function ReadINIValue( _
ByVal SectionName As String, ByVal KeyName As String, _
ByVal DefaultValue As String) As String
Dim sBuff As String
Dim X As Long

    sBuff = Space$(1024)
    X = GetPrivateProfileString(StrPtr(SectionName), StrPtr(KeyName), StrPtr(DefaultValue), _
    StrPtr(sBuff), Len(sBuff), StrPtr(m_INIPath))
    ReadINIValue = Left$(sBuff, X)
End Function

Public Sub WriteINIValue( _
ByVal SectionName As String, ByVal KeyName As String, _
ByVal KeyValue As String)

Dim X As Long
    X = WritePrivateProfileString(StrPtr(SectionName), StrPtr(KeyName), StrPtr(KeyValue), StrPtr(m_INIPath))
End Sub



