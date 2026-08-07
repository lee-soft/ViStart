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
Option Explicit

Public ImageName As String

'Set at build time when the AutomaticDestinations blobs hold entries for this
'app — those items only materialize in JumpListDrawer, so IsEmpty must know
'about them or the jumplist arrow never shows.
Public HasAutoDestinationItems As Boolean

Private m_mruKey As Collection

Public Function AddMRURegKey(thisMRUKey As RegistryKey)
    m_mruKey.Add thisMRUKey
End Function

Public Function IsEmpty() As Boolean
    IsEmpty = True

    If m_mruKey.count > 0 Or HasAutoDestinationItems Then
        IsEmpty = False
    End If
End Function

Public Function GetMRUList()

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

