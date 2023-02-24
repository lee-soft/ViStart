VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Enum RegistryHive
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum RegistryValueKind
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_MULTI_SZ = 7
    REG_QWORD = 11
End Enum

Private m_hive As RegistryHive

Property Let Hive(ByVal newHive As RegistryHive)
    m_hive = newHive
End Property

Public Function SetValue(ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Variant, Optional ByVal ValueType As RegistryValueKind = REG_SZ) As Boolean
    Dim hKey As Long
    Dim Result As Long
    
    Result = RegCreateKeyEx(m_hive, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0&, hKey, 0&)
    If Result = ERROR_SUCCESS Then
        Select Case ValueType
            Case REG_SZ, REG_EXPAND_SZ
                Result = RegSetValueEx(hKey, ValueName, 0&, ValueType, Value, Len(Value) + 1)
            Case REG_BINARY, REG_DWORD, REG_QWORD
                Result = RegSetValueEx(hKey, ValueName, 0&, ValueType, Value, LenB(Value))
            Case REG_MULTI_SZ
                Dim MultiValue() As String
                MultiValue = Split(Value, vbNullChar)
                Result = RegSetValueEx(hKey, ValueName, 0&, ValueType, MultiValue(0), Len(Value) + 2)
        End Select
        RegCloseKey hKey
    End If
    
    SetValue = (Result = ERROR_SUCCESS)
End Function

Public Function GetValue(ByVal KeyName As String, ByVal ValueName As String, Optional ByVal DefaultValue As Variant = vbNull) As Variant

    Dim hKey As Long
    Dim Result As Long
        
    Result = RegOpenKeyEx(m_hive, KeyName, 0&, KEY_READ, hKey)
    If Result <> ERROR_SUCCESS Then
        ' Failed to open key
        GetValue = DefaultValue
        Exit Function
    End If
    
    Dim ValueType As Long
    Dim ValueDataBuffer() As Byte
    Dim ValueData As String
    Dim ValueSize As Long
    
    ReDim ValueDataBuffer(0)
    Result = RegQueryValueEx(hKey, ValueName, 0&, ValueType, ValueDataBuffer(0), ValueSize)
    
    ReDim ValueDataBuffer(0 To ValueSize)
    Result = RegQueryValueEx(hKey, ValueName, 0&, ValueType, ValueDataBuffer(0), ValueSize)
    
    ValueData = StrConv(ValueDataBuffer, vbUnicode)
    
    If Result <> ERROR_SUCCESS Then
        RegCloseKey hKey
        GetValue = DefaultValue
        Exit Function
    End If
    
    Select Case ValueType
    
    Case REG_SZ, REG_EXPAND_SZ
        GetValue = Left$(ValueData, ValueSize - 1)
    Case REG_BINARY, REG_DWORD, REG_QWORD
        GetValue = ValueDataBuffer
    Case REG_MULTI_SZ
        Dim unformattedData As String
        unformattedData = Left$(ValueData, ValueSize - 1)
        
        GetValue = Replace(unformattedData, ChrW(0), vbCrLf)
        
    End Select
    
    RegCloseKey hKey
End Function
