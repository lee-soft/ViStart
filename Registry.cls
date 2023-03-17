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
'--------------------------------------------------------------------------------
'    Component  : Registry
'    Project    : ViStart
'    Author     : Lee Chantrey, lchantrey@gmail.com
'
'    Description: Enables the consumer to use Registry.CurrentUser/LocalMachine
'                 Like in .NET, So you can do this
'                 Registry.CurrentUser.OpenSubKey etc
'--------------------------------------------------------------------------------


Option Explicit

Private m_currentUser As RegistryKey
Private m_localMachine As RegistryKey
Private m_classesRoot As RegistryKey
Private m_users As RegistryKey

Public Property Get CurrentUser() As RegistryKey
    Set CurrentUser = m_currentUser
End Property
    
Public Property Get LocalMachine() As RegistryKey
    Set LocalMachine = m_localMachine
End Property

Public Property Get ClassesRoot() As RegistryKey
    Set ClassesRoot = m_classesRoot
End Property

Public Property Get Users() As RegistryKey
    Set Users = m_users
End Property

Private Sub Class_Initialize()
    Set m_currentUser = New RegistryKey
    m_currentUser.Hive = HKEY_CURRENT_USER
    
    Set m_localMachine = New RegistryKey
    m_localMachine.Hive = HKEY_LOCAL_MACHINE
    
    Set m_classesRoot = New RegistryKey
    m_classesRoot.Hive = HKEY_CLASSES_ROOT
    
    Set m_users = New RegistryKey
    m_users.Hive = HKEY_USERS
End Sub
