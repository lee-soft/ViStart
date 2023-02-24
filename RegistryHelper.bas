Attribute VB_Name = "RegistryHelper"
Attribute VB_PredeclaredId = True

Option Explicit

'Marked as Default member:
Public Property Get Registry() As RegistryKey
Attribute Setting.VB_UserMemId = 0
Attribute Setting.VB_MemberFlags = "200"
    Set Registry = New RegistryKey
End Property

