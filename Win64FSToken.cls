VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win64FSToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_hOldValue As Long
Private m_disabled As Boolean

Public Property Get Disabled()
    Disabled = m_disabled
End Property

Public Function EnableFS()
    If m_disabled Then
        m_disabled = False
        
        Wow64RevertWow64FsRedirection m_hOldValue
    End If
End Function

Private Sub Class_Initialize()
    m_disabled = False
    m_hOldValue = 0

    If Wow64DisableWow64FsRedirection(m_hOldValue) <> 0 Then
        m_disabled = True
    End If
End Sub

Private Sub Class_Terminate()
    EnableFS
End Sub
