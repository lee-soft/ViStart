VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShellReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_RegObj As New clsRegistry
Private WshShell As Object

Public Property Get RegObj() As clsRegistry
    Set RegObj = m_RegObj
End Property

Function WriteTo(sPath As String, varData As Variant)
    
Dim sKeyName As String
    
    sPath = sCon_Reg_AppPath & sPath
    
    sKeyName = StrEnd(sPath, "\")
    sPath = Mid$(sPath, 1, Len(sPath) - Len(sKeyName))
    
    m_RegObj.CreateKey HKEY_CURRENT_USER, sPath
    
    If VarType(varData) = vbBoolean Then
        m_RegObj.SetBooleanValue HKEY_CURRENT_USER, sPath, sKeyName, CBol(varData)
    ElseIf VarType(varData) = vbLong Then
        m_RegObj.SetLongValue HKEY_CURRENT_USER, sPath, sKeyName, CLng(varData)
    Else
        m_RegObj.SetStringValue HKEY_CURRENT_USER, sPath, sKeyName, CStr(varData)
    End If
    
End Function

Function GetAppSettingBooleon(sPath As String, Optional Default) As Boolean
    
Dim sKeyName As String
    
    sPath = sCon_Reg_AppPath & sPath
    
    sKeyName = StrEnd(sPath, "\")
    sPath = Mid$(sPath, 1, Len(sPath) - Len(sKeyName))
    
    GetAppSettingBooleon = m_RegObj.GetBooleanValue(HKEY_CURRENT_USER, sPath, sKeyName, CBol(Default))

End Function

Function GetAppSettingLong(sPath As String, Optional Default) As Long
    
Dim sKeyName As String
    
    sPath = sCon_Reg_AppPath & sPath
    
    sKeyName = StrEnd(sPath, "\")
    sPath = Mid$(sPath, 1, Len(sPath) - Len(sKeyName))
    
    GetAppSettingLong = m_RegObj.GetLongValue(HKEY_CURRENT_USER, sPath, sKeyName, CLng(Default))

End Function

Function GetAppSetting(sPath As String, Optional Default)
    
Dim sKeyName As String
    
    sPath = sCon_Reg_AppPath & sPath
    
    sKeyName = StrEnd(sPath, "\")
    sPath = Mid$(sPath, 1, Len(sPath) - Len(sKeyName))
    
    GetAppSetting = m_RegObj.GetStringValue(HKEY_CURRENT_USER, sPath, sKeyName, CStr(Default))

End Function

Function RegWrite(sPath As String, Data) As Boolean

    On Error GoTo Handler

    RegWrite = WshShell.RegWrite(sPath, Data)
    RegWrite = True
    
    Exit Function
    
Handler:
    MsgBox Err.Description, vbCritical, "Registry::Write - Error"
    RegWrite = False
    
End Function

Function Read(sPath As String, Optional Default, Optional bIgnoreError As Boolean = False)

    On Error GoTo Handler

    Read = WshShell.RegRead(sPath)

    Exit Function
    
Handler:
    If Err.Number = -2147024894 Then
        'Invalid Reg Key
        
        If IsMissing(Default) Then
            If Not bIgnoreError Then
                MessageBox 0, Err.Description, "Registry::Read - Error", vbCritical
            End If
            
            Read = ""
        Else
            Read = Default
        End If
    Else
        MessageBox 0, Err.Description, "Registry::Object - Error", vbCritical
        Read = ""
    End If

End Function

Function Delete(sPath As String) As Boolean

    On Error GoTo Handler

    WshShell.RegDelete sPath
    Delete = True
    
    Exit Function
    
Handler:
    Delete = False
    'MsgBox Err.description, vbCritical, "Registry::Delete - Error"
    'Delete = ""

End Function

Function RegReadBinaryA(sPath As String) As String
    'Returns a string, when registry is stored as ansi
    '(Char) (Char)

Dim lngByteIndex As Long
Dim bByteArr As Variant

Dim sNewString As String
    
    On Error GoTo Handler
    
    bByteArr = Registry.Read(sPath, "")

    'Convert to string
    For lngByteIndex = 0 To UBound(bByteArr)
        sNewString = sNewString & Chr$(bByteArr(lngByteIndex))
    Next
    
    RegReadBinaryA = sNewString
    
    Exit Function
Handler:

End Function

Function RegReadBinaryW(sPath As String) As String
    'Returns a string, when registry is stored as unicode
    '00 (Char) 00 (Char) etc

Dim lngByteIndex As Long
Dim bByteArr As Variant

Dim sNewString As String
    
    On Error GoTo Handler
    
    bByteArr = Registry.Read(sPath, "")

    'Convert to string
    For lngByteIndex = 0 To UBound(bByteArr)
        sNewString = sNewString & ChrB$(bByteArr(lngByteIndex))
    Next
    
    RegReadBinaryW = sNewString
    
    Exit Function
Handler:

End Function

Private Sub Class_Initialize()
    Set WshShell = CreateObject("WScript.Shell")
End Sub

Private Sub Class_Terminate()
    Set WshShell = Nothing
End Sub
