VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NO_MORE_ITEMS       As Long = 259&

Private Const KEY_WOW64_32KEY = &H200
Private Const KEY_WOW64_64KEY = &H100

Public Enum REG_HKEY
        HKEY_CLASSES_ROOT = &H80000000
        HKEY_CURRENT_CONFIG = &H80000005
        HKEY_CURRENT_USER = &H80000001
        HKEY_DYN_DATA = &H80000006
        HKEY_LOCAL_MACHINE = &H80000002
        HKEY_PERFORMANCE_DATA = &H80000004
        HKEY_USERS = &H80000003
End Enum

Public Enum REG_TYPE
       NONE = 0                       ' No value type
       sz = 1                         ' Unicode nul terminated string
       EXPAND_SZ = 2                  ' Unicode nul terminated string
       Binary = 3                     ' Free form binary
       DWORD = 4                      ' 32-bit number
       DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
       DWORD_BIG_ENDIAN = 5           ' 32-bit number
       LINK = 6                       ' Symbolic Link (unicode)
       MULTI_SZ = 7                   ' Multiple Unicode strings
End Enum

Public Function KeyExists(hKey As REG_HKEY, KeyName As String) As Boolean

Dim RegKeyHandle As Long

    If RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle) = ERROR_SUCCESS Then
        RegCloseKey RegKeyHandle 'close the key
        
        KeyExists = True
    End If

End Function

Public Function CreateKey(hKey As REG_HKEY, KeyName As String, Optional Access64Bit As Boolean = False) As Boolean
Dim RegKeyHandle As Long
Dim ret As Long

If Access64Bit And Is64bit Then
    ret = RegCreateKeyEx(CLng(hKey), KeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, ByVal 0&, RegKeyHandle, ByVal 0&)
Else
    ret = RegCreateKeyEx(CLng(hKey), KeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, RegKeyHandle, ByVal 0&)
End If

If ret <> ERROR_SUCCESS Then
   RegCloseKey RegKeyHandle
   CreateKey = False
Else
   RegCloseKey RegKeyHandle
   CreateKey = True
End If
End Function

'EnumerateKeys
Public Function EnumKeys(hKey As REG_HKEY, ByVal sPath As String, Key() As String) As Long
    
Dim sName   As String
Dim RetVal  As Long
Dim RegKeyHandle As Long
    
    Erase Key 'clear array
    
    'try to open key
    If RegOpenKeyEx(hKey, sPath, 0, KEY_ALL_ACCESS, RegKeyHandle) = ERROR_SUCCESS Then

        EnumKeys = 0 'array is 0-based
        sName = Space$(255)
        RetVal = Len(sName)
        
        'do while not we get ERROR_NO_MORE_ITEMS
        While RegEnumKeyEx(RegKeyHandle, EnumKeys, sName, RetVal, ByVal 0&, _
                           vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            
            ReDim Preserve Key(EnumKeys) 'incermenting array (+1)
            
            Key(EnumKeys) = Left$(sName, RetVal) 'adding key name to array
                        
            'preparing values for next data
            EnumKeys = EnumKeys + 1 'incerment the counter
            sName = Space$(255)
            RetVal = Len(sName)
            
        Wend
    
        RegCloseKey RegKeyHandle 'close the key
    Else
        EnumKeys = -1 'error (key doesn't exists)
    End If
    
End Function

Public Function DeleteEmptyKey(hKey As REG_HKEY, KeyName As String) As Boolean
Dim ret As Long
ret = RegDeleteKey(CLng(hKey), KeyName)
If ret <> ERROR_SUCCESS Then
   DeleteEmptyKey = False
Else
   DeleteEmptyKey = True
End If
End Function

Public Function DeleteKey(hKey As REG_HKEY, KeyName As String) As Boolean
Dim ret As Long
ret = SHDeleteKey(CLng(hKey), KeyName)
If ret <> ERROR_SUCCESS Then
   DeleteKey = False
Else
   DeleteKey = True
End If
End Function

Public Function SetStringValue(hKey As REG_HKEY, KeyName As String, ValueName As String, Value As String, Optional Access64Bit As Boolean = False) As Boolean
Dim ret As Long
Dim RegKeyHandle As Long

If Access64Bit And Is64bit Then
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, RegKeyHandle)
Else
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
End If

If ret <> ERROR_SUCCESS Then
   RegCloseKey RegKeyHandle
   SetStringValue = False
   Exit Function
End If
ret = RegSetValueEx(RegKeyHandle, StrPtr(ValueName), 0, REG_SZ, ByVal StrPtr(Value), LenB(Value))
RegCloseKey RegKeyHandle
SetStringValue = (ret = ERROR_SUCCESS)
End Function

Public Function SetLongValue(hKey As REG_HKEY, KeyName As String, ValueName As String, Value As Long, Optional Access64Bit As Boolean = False) As Boolean
Dim ret As Long
Dim RegKeyHandle As Long

If Access64Bit And Is64bit Then
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, RegKeyHandle)
Else
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
End If

If ret <> ERROR_SUCCESS Then
   RegCloseKey RegKeyHandle
   SetLongValue = False
   Exit Function
End If
ret = RegSetValueEx(RegKeyHandle, StrPtr(ValueName), ByVal 0&, REG_DWORD, CLng(Value), Len(Value))
RegCloseKey RegKeyHandle
SetLongValue = (ret = ERROR_SUCCESS)
End Function


Public Function GetStringValue(hKey As REG_HKEY, KeyName As String, KeyValue As String, Optional Default As String = "", Optional Access64Bit As Boolean = False) As String
Dim ret As Long
Dim RegKeyHandle As Long
Dim Buffer(0 To 16299) As Byte
Dim BufferLen As Long
Dim i As Long

If Access64Bit And Is64bit Then
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, RegKeyHandle)
Else
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
End If

If ret <> ERROR_SUCCESS Then
   RegCloseKey RegKeyHandle
   GetStringValue = Default
   Exit Function
End If
BufferLen = 16300
ret = RegQueryValueEx(RegKeyHandle, StrPtr(KeyValue), ByVal 0&, REG_SZ, Buffer(0), BufferLen)
If ret <> ERROR_SUCCESS Then
   GetStringValue = Default
   Exit Function
End If

GetStringValue = StripTerminator(Buffer)
End Function

Public Function GetLongValue(hKey As REG_HKEY, KeyName As String, KeyValue As String, Optional Default As Long = -1, Optional Access64Bit As Boolean = False) As Long
Dim ret As Long
Dim RegKeyHandle As Long
Dim Buffer As Long
Dim BufferLen As Long

Dim i As Long

If Access64Bit And Is64bit Then
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, RegKeyHandle)
Else
    ret = RegOpenKeyEx(CLng(hKey), KeyName, 0, KEY_ALL_ACCESS, RegKeyHandle)
End If

If ret <> ERROR_SUCCESS Then
   RegCloseKey RegKeyHandle
   GetLongValue = Default
   Exit Function
End If
BufferLen = Len(Buffer)
ret = RegQueryValueEx(RegKeyHandle, StrPtr(KeyValue), ByVal 0&, REG_SZ, Buffer, BufferLen)
If ret <> ERROR_SUCCESS Then
   GetLongValue = Default
   Exit Function
End If
GetLongValue = Buffer
End Function

Public Function SetBooleanValue(hKey As REG_HKEY, KeyName As String, KeyValue As String, Value As Boolean) As Boolean
   SetBooleanValue = SetLongValue(hKey, KeyName, KeyValue, IIf(Value = False, 0, 1), True)
End Function

Public Function GetBooleanValue(hKey As REG_HKEY, KeyName As String, KeyValue As String, Optional Default As Boolean = False) As Boolean
Dim v As Long
Dim ret As Long
v = GetLongValue(hKey, KeyName, KeyValue, -1, True)

If v <> -1 Then
   GetBooleanValue = IIf(v = 0, False, True)
   Exit Function
End If
GetBooleanValue = Default
End Function

Public Function Read(ByVal thePath As String, Optional ByVal DefaultValue As String) As String
    
Dim theStart As String
Dim theKeyName As String

    theStart = UCase$(Split(thePath, "\")(0))
    
    thePath = Split(thePath, "\", 2)(1)
    theKeyName = Mid$(thePath, InStrRev(thePath, "\") + 1)
    thePath = Mid$(thePath, 1, Len(thePath) - Len(theKeyName) - 1)
    
    If theStart = "HKCU" Then
        Read = GetStringValue(HKEY_CURRENT_USER, thePath, theKeyName, DefaultValue)
    ElseIf theStart = "HKLM" Then
        Read = GetStringValue(HKEY_LOCAL_MACHINE, thePath, theKeyName, DefaultValue)
    End If
    
End Function


