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
'--------------------------------------------------------------------------------
'    Component  : RegistryKey
'    Project    : ViStart
'    Author     : Lee Chantrey, lchantrey@gmail.com
'
'    Description: Port of the .NET RegistryKey class
'--------------------------------------------------------------------------------

Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_READ As Long = &H20019
Private Const KEY_WRITE As Long = &H20006
Private Const KEY_ALL_ACCESS As Long = &HF003F

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

Private m_hKey As Long
Private m_hKeySet As Boolean
Private m_Path As String
Private m_supressErrors As Boolean

Property Let Hive(ByVal newHive As RegistryHive)
    m_hive = newHive
End Property

Property Let Key(ByVal NewKey As Long)
    m_hKey = NewKey
    m_hKeySet = True
End Property

Property Get Key() As Long
    If m_hKeySet Then
        Key = m_hKey
    Else
        Key = m_hive
    End If
End Property

Property Let Path(ByVal newPath As String)
    m_Path = newPath
End Property

Property Get Path() As String
    Path = m_Path
End Property

Public Sub DeleteValue(ByVal Name As String)
    RegDeleteValue Me.Key, Name
End Sub

Public Function GetSubKeyNames() As Collection

Dim subkey_num As Integer
Dim subkey_name As String

Dim Length As Long

    Dim SubKeysNames As Collection
    Set SubKeysNames = New Collection
    Set GetSubKeyNames = SubKeysNames

    ' Enumerate the subkeys.
    subkey_num = 0

    Do
        ' Enumerate subkeys until we get an error.
        Length = 1024
        subkey_name = Space$(Length)
        If RegEnumKey(Me.Key, subkey_num, _
            subkey_name, Length) _
                <> ERROR_SUCCESS Then Exit Do
                
        subkey_num = subkey_num + 1
        subkey_name = Left$(subkey_name, InStr(subkey_name, Chr$(0)) - 1)
        
        SubKeysNames.Add subkey_name
    Loop

End Function

Public Function CreateSubKey(ByVal subKeyName As String) As RegistryKey
    Dim hKey As Long
    Dim subKeyPath As String
    Dim returnObject As RegistryKey
    
    ' Combine the key path and subkey name to create the full subkey path
    
    If Me.Path <> vbNullString Then
        subKeyPath = Me.Path & "\" & subKeyName
    Else
        subKeyPath = subKeyName
    End If
    
    ' Attempt to create the subkey
    Dim result As Long
    result = RegCreateKeyEx(m_hive, subKeyPath, 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, ByVal 0&)

    ' Check for success
    If Not result = ERROR_SUCCESS Then
        ' Failed to create subkey
        
        LogError "RegistryKey::Unable to create subkey: " & subKeyName
        If Not m_supressErrors Then Err.Raise result, "RegistryKey", "Unable to create subkey"
        Exit Function
    End If
    
    ' Subkey was created successfully
    RegCloseKey hKey
    
    Set returnObject = Me.OpenSubKey(subKeyName, True)
    Set CreateSubKey = returnObject
End Function

Public Function OpenSubKey(ByVal subKeyName As String, Optional ByVal writable As Boolean = False) As RegistryKey
    Dim regKey As RegistryKey: Set regKey = New RegistryKey
    regKey.Hive = m_hive
    
    Dim access As Long
    Dim hSubKey As Long
    Dim result As Long
    
    ' Determine the access level
    If writable Then
        access = KEY_READ Or KEY_WRITE
    Else
        access = KEY_READ
    End If
    
    result = RegOpenKeyEx(Me.Key, subKeyName, 0, access, hSubKey)
    
    ' Open the subkey
    If Not result = ERROR_SUCCESS Then
        LogError "RegistryKey::Unable to open sub key: " & subKeyName
        If Not m_supressErrors Then Err.Raise result, "RegistryKey", "Unable to open Registry sub key: " & subKeyName
        Exit Function
    End If
    
    regKey.Key = hSubKey
    regKey.Path = subKeyName
    
    Set OpenSubKey = regKey
End Function

Public Function SetValue(ByVal ValueName As String, ByVal Value As Variant) As Boolean
    Dim hKey As Long
    Dim result As Long
    Dim valueAsBytes() As Byte
    
    Dim ValueType As RegistryValueKind
    Dim vbValueType As Long: vbValueType = VarType(Value)
    
    If vbValueType = vbString Then
        ValueType = REG_SZ
    ElseIf vbValueType = vbLong Or vbValueType = vbInteger Then
        ValueType = REG_DWORD
    Else
        LogError "RegistryKey, unimplented type: " & vbValueType
        If Not m_supressErrors Then Err.Raise 0, "RegistryKey", "Unimplmented type!"
    End If
    
    result = RegOpenKeyEx(Me.Key, "", 0, KEY_ALL_ACCESS, hKey)
    If result = ERROR_SUCCESS Then
    
        Select Case ValueType
            Case REG_SZ, REG_EXPAND_SZ
                valueAsBytes = StrConv(Value, vbFromUnicode)
                
            Case REG_BINARY, REG_DWORD, REG_QWORD
                valueAsBytes = LongToDWordByteArray(Value)
                
            'Case REG_MULTI_SZ
            '    Dim MultiValue() As String
            '    MultiValue = Split(value, vbNullChar)
            '    result = RegSetValueEx(Me.Key, ValueName, 0&, ValueType, MultiValue(0), Len(value) + 2)
                
        End Select
        
        result = RegSetValueEx(hKey, ValueName, 0&, ValueType, valueAsBytes(0), UBound(valueAsBytes) + 1)
        RegCloseKey hKey
    End If
    
    SetValue = (result = ERROR_SUCCESS)
End Function

Public Function GetValue(ByVal ValueName As String, Optional ByVal DefaultValue As Variant = vbNull) As Variant

    Dim hKey As Long
    Dim result As Long
        
    result = RegOpenKeyEx(Me.Key, "", 0&, KEY_READ, hKey)
    If result <> ERROR_SUCCESS Then
        LogError "RegistryKey", "Unable to open key: & ValueName"
        If Not m_supressErrors Then Err.Raise result, "RegistryKey", "Unable to open key!"
        ' Failed to open key
        GetValue = DefaultValue
        Exit Function
    End If
    
    Dim ValueType As Long
    Dim ValueDataBuffer() As Byte
    Dim ValueData As String
    Dim ValueSize As Long
    
    ReDim ValueDataBuffer(0)
    result = RegQueryValueEx(hKey, ValueName, 0&, ValueType, ValueDataBuffer(0), ValueSize)
    
    ReDim ValueDataBuffer(0 To ValueSize)
    result = RegQueryValueEx(hKey, ValueName, 0&, ValueType, ValueDataBuffer(0), ValueSize)
    
    ValueData = StrConv(ValueDataBuffer, vbUnicode)
    
    If result <> ERROR_SUCCESS Then
        RegCloseKey hKey
        GetValue = DefaultValue
        
        Debug.Print "RegistryKey", "Unable to read key: " & ValueName
        If Not m_supressErrors Then Err.Raise result, "RegistryKey", "Unable to read key: " & ValueName
        Exit Function
    End If
    
    Select Case ValueType
    
    Case REG_SZ, REG_EXPAND_SZ
        GetValue = Left$(ValueData, ValueSize - 1)
    Case REG_BINARY, REG_QWORD
        GetValue = ValueDataBuffer
    Case REG_DWORD
        GetValue = ConvertBytesToLong(ValueDataBuffer)
        
    Case REG_MULTI_SZ
        Dim unformattedData As String
        unformattedData = Left$(ValueData, ValueSize - 1)
        
        GetValue = Replace(unformattedData, ChrW$(0), vbCrLf)
        
    End Select
    
    RegCloseKey hKey
End Function

Private Sub Class_Initialize()
    m_supressErrors = True
End Sub

Private Sub Class_Terminate()
    If m_hKeySet Then RegCloseKey m_hKey
End Sub
