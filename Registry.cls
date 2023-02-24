VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get RegObject() As RegistryKey
Attribute RegObject.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RegObject
    Set RegObject = New RegistryKey
End Property

Public Property Get CurrentUser() As RegistryKey
    Set CurrentUser = New RegistryKey
    CurrentUser.Hive = HKEY_CURRENT_USER
End Property
    
Public Property Get LocalMachine() As RegistryKey
    Set LocalMachine = New RegistryKey
    LocalMachine.Hive = HKEY_LOCAL_MACHINE
End Property

Public Property Get ClassesRoot() As RegistryKey
    Set ClassesRoot = New RegistryKey
    LocalMachine.Hive = HKEY_CLASSES_ROOT
End Property
