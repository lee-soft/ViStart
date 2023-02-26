VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JumpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public ImageName As String

Private m_mruKey As Collection

Public Function AddMRURegKey(thisMRUKey As RegistryKey)
    m_mruKey.Add thisMRUKey
End Function

Public Function IsEmpty() As Boolean
    IsEmpty = True
    
    If m_mruKey.count > 0 Then
        IsEmpty = False
    End If
End Function

Public Function GetMRUList()
    Debug.Print "GetMRUList() " & ImageName

Dim thisMRUKey As RegistryKey
Dim m_lnkCache()
 
    For Each thisMRUKey In m_mruKey
        m_lnkCache = ConcatArray(m_lnkCache, GetMRUListForKey(thisMRUKey))
    Next
    
    GetMRUList = m_lnkCache

End Function

Private Sub Class_Initialize()
    Set m_mruKey = New Collection
End Sub

