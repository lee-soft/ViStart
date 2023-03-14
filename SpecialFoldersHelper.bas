Attribute VB_Name = "SpecialFoldersHelper"
Option Explicit

Public Sub RestoreDefaultFolders()

Dim myDocs As String: myDocs = "My Documents"
Dim myPics As String: myPics = "My Pictures"
Dim myMus As String: myMus = "My Music"
Dim myVid As String: myVid = "My Video"
Dim downloads As String: downloads = "Downloads"
Dim games As String: games = "Games"
Dim ctrPanel As String: ctrPanel = "Control Panel"
Dim objects3D As String: objects3D = "3D Objects"
Dim desk As String: desk = "Desktop"
Dim userName As String

    With Registry.ClassesRoot.OpenBaseKey(HKEY_CLASSES_ROOT)
        .DeleteSubKey "CLSID\{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}"
        .DeleteSubKey "CLSID\{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}"
    
        .DeleteSubKey "CLSID\{1CF1260C-4DD0-4EBB-811F-33C572699FDE}"
        .DeleteSubKey "CLSID\{A0953C92-50DC-43BF-BE83-3742FED03C9C}"
        .DeleteSubKey "CLSID\{374DE290-123F-4565-9164-39C4925E467B}"
    
        .DeleteSubKey "CLSID\{d3162b92-9365-467a-956b-92703aca08af}"
        .DeleteSubKey "CLSID\{24ad3ad4-a569-4530-98e1-ab02f9417aa8}"
        .DeleteSubKey "CLSID\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}"
    End With
    
    userName = Environment.GetEnvironmentVariable("UserName")

Dim Key As RegistryKey: Set Key = RegistryKey.OpenBaseKey(HKEY_LOCAL_MACHINE). _
                                              CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension")
                                                                                
    If Key Is Nothing Then
        MsgBox "Could not open key, aborting!"
        Exit Sub
    End If
    
    Key.SetValue "Type", "group", REG_SZ
    Key.SetValue "Text", "Special Folders", REG_SZ
    Key.SetValue "Bitmap", "%SystemRoot%\system32\SHELL32.dll,211", REG_EXPAND_SZ

    'Windows XP,
    If (Environment.OSVersion.Major = 5 And Environment.OSVersion.Minor = 1) Then
        FixFolderPaths
        
        Call CreateCLSID("{59031A47-3F72-44A7-89C5-5595FE6B30EE}", "%UserProfile%", "%SystemRoot%\system32\SHELL32.dll,-269", userName, "", 1)
        Call CreateFolderOption("{59031A47-3F72-44A7-89C5-5595FE6B30EE}", "%SystemRoot%\system32\SHELL32.dll,-269", userName)
        
        Call CreateCLSID("{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}", "%UserProfile%\" & myDocs, "%SystemRoot%\system32\SHELL32.dll,-235", myDocs, "%SystemRoot%\system32\SHELL32.dll,-30349", 2)
        Call CreateFolderOption("{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}", "%SystemRoot%\system32\SHELL32.dll,-235", myDocs)
        
        Call CreateCLSID("{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}", "%UserProfile%\" & myPics, "%SystemRoot%\system32\SHELL32.dll,-236", myPics, "%SystemRoot%\system32\SHELL32.dll,-12688", 3)
        Call CreateFolderOption("{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}", "%SystemRoot%\system32\SHELL32.dll,-236", myPics)
        
        Call CreateCLSID("{1CF1260C-4DD0-4EBB-811F-33C572699FDE}", "%UserProfile%\" & myMus, "%SystemRoot%\system32\SHELL32.dll,-237", myMus, "%SystemRoot%\system32\SHELL32.dll,-12689", 4)
        Call CreateFolderOption("{1CF1260C-4DD0-4EBB-811F-33C572699FDE}", "%SystemRoot%\system32\SHELL32.dll,-237", myMus)
        
        Call CreateCLSID("{A0953C92-50DC-43BF-BE83-3742FED03C9C}", "%UserProfile%\" & myVid, "%SystemRoot%\system32\SHELL32.dll,-238", myVid, "%SystemRoot%\system32\SHELL32.dll,-12690", 5)
        Call CreateFolderOption("{A0953C92-50DC-43BF-BE83-3742FED03C9C}", "%SystemRoot%\system32\SHELL32.dll,-238", myVid)
        
        Call CreateCLSID("{374DE290-123F-4565-9164-39C4925E467B}", "%UserProfile%\" & downloads, "%SystemRoot%\system32\SHELL32.dll,-14", downloads, "", 0)
        Call CreateCLSID("{374DE290-123F-4565-9164-39C4925E467B}", "%UserProfile%\" & downloads, "%SystemRoot%\system32\inetcpl.cpl,-4460", downloads, "", 6)

        Call CreateFolderOption("{374DE290-123F-4565-9164-39C4925E467B}", "%SystemRoot%\system32\SHELL32.dll,-14", downloads)
        Call CreateFolderOption("{374DE290-123F-4565-9164-39C4925E467B}", "%SystemRoot%\system32\ieframe.dll,-113", downloads)

        Call CreateCLSID("{ED228FDF-9EA8-4870-83b1-96b02CFE0D52}", Environ("ALLUSERSPROFILE") & "\Start Menu\Programs\Games", "%SystemRoot%\system32\xpsp3res.dll,-100", games, "", 7)
        Call CreateFolderOption("{ED228FDF-9EA8-4870-83b1-96b02CFE0D52}", "%SystemRoot%\system32\xpsp3res.dll,-100", games)
        
        Call CreateFolderOption("{21EC2020-3AEA-1069-A2DD-08002B30309D}", "%SystemRoot%\system32\SHELL32.dll,-22", ctrPanel)
        Dim showDesktopKey As RegistryKey: Set showDesktopKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension\{21EC2020-3AEA-1069-A2DD-08002B30309D}", True)
        If Not (showDesktopKey Is Nothing) Then
            showDesktopKey.DeleteValue "SHOWONDESKTOP"
            Set showDesktopKey = Nothing
        End If
        
    'Windows 10/11
    ElseIf (Environment.OSVersion.Major = 10) Then
    
        Call CreateFolderOption("{59031A47-3F72-44A7-89C5-5595FE6B30EE}", "%SystemRoot%\system32\imageres.dll,-123", userName)
        Call CreateFolderOption("{d3162b92-9365-467a-956b-92703aca08af}", "%SystemRoot%\system32\imageres.dll,-112", myDocs)
        Call CreateFolderOption("{24ad3ad4-a569-4530-98e1-ab02f9417aa8}", "%SystemRoot%\system32\imageres.dll,-113", myPics)
        Call CreateFolderOption("{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}", "%SystemRoot%\system32\imageres.dll,-108", myMus)
        Call CreateFolderOption("{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}", "%SystemRoot%\system32\imageres.dll,-189", myVid)
        Call CreateFolderOption("{088e3905-0323-4b02-9826-5d99428e115f}", "%SystemRoot%\system32\imageres.dll,-184", downloads)
        
        Call CreateFolderOption("{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}", "%SystemRoot%\system32\imageres.dll,-198", objects3D)
        Registry.LocalMachine.OpenBaseKey(HKEY_LOCAL_MACHINE).DeleteSubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}\SHOWONDESKTOP"
        
        Call CreateFolderOption("{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}", "%SystemRoot%\system32\SHELL32.dll,-22", ctrPanel)
        Registry.LocalMachine.OpenBaseKey(HKEY_LOCAL_MACHINE).DeleteSubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension\{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}\SHOWONDESKTOP"
        
    
    'Windows Vista or 8
    ElseIf (Environment.OSVersion.Major = 6 And Environment.OSVersion.Minor = 0) Or _
           (Environment.OSVersion.Major = 6 And Environment.OSVersion.Minor = 3) Then
           
        Call CreateCLSID("{59031A47-3F72-44A7-89C5-5595FE6B30EE}", "%UserProfile%", "%SystemRoot%\system32\imageres.dll,-123", userName, "", "")
        Call CreateFolderOption("{59031A47-3F72-44A7-89C5-5595FE6B30EE}", "%SystemRoot%\system32\imageres.dll,-123", userName)
        
        Call CreateCLSID("{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}", "%UserProfile%" & "\" & myDocs, "%SystemRoot%\system32\imageres.dll,-112", myDocs, "%SystemRoot%\system32\SHELL32.dll,-30349")
        Call CreateFolderOption("{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}", "%SystemRoot%\system32\imageres.dll,-112", myDocs)
        
        Call CreateCLSID("{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}", "%UserProfile%" & "\" & myPics, "%SystemRoot%\system32\imageres.dll,-113", myPics, "%SystemRoot%\system32\SHELL32.dll,-12688")
        Call CreateFolderOption("{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}", "%SystemRoot%\system32\imageres.dll,-113", myPics)
        
        Call CreateCLSID("{1CF1260C-4DD0-4EBB-811F-33C572699FDE}", "%UserProfile%" & "\" & myMus, "%SystemRoot%\system32\imageres.dll,-108", myMus, "%SystemRoot%\system32\SHELL32.dll,-12689")
        Call CreateFolderOption("{1CF1260C-4DD0-4EBB-811F-33C572699FDE}", "%SystemRoot%\system32\imageres.dll,-108", myMus)
        
        Call CreateCLSID("{A0953C92-50DC-43BF-BE83-3742FED03C9C}", "%UserProfile%" & "\" & myVid, "%SystemRoot%\system32\imageres.dll,-189", myVid, "%SystemRoot%\system32\SHELL32.dll,-12690")
        Call CreateFolderOption("{A0953C92-50DC-43BF-BE83-3742FED03C9C}", "%SystemRoot%\system32\imageres.dll,-189", myVid)
    
        Call CreateCLSID("{374DE290-123F-4565-9164-39C4925E467B}", "%UserProfile%" & "\" & downloads, "%SystemRoot%\system32\imageres.dll,-189", downloads, "%SystemRoot%\system32\SHELL32.dll,-12690")
        Call CreateFolderOption("{374DE290-123F-4565-9164-39C4925E467B}", "%SystemRoot%\system32\imageres.dll,-189", downloads)
    
        CreateFolderOption "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}", "%SystemRoot%\system32\imageres.dll,-183", desk
        Registry.LocalMachine.OpenBaseKey(HKEY_LOCAL_MACHINE).DeleteSubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}\SHOWONDESKTOP"
        
        CreateFolderOption "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}", "%SystemRoot%\system32\SHELL32.dll,-22", objects3D
        Registry.LocalMachine.OpenBaseKey(HKEY_LOCAL_MACHINE).DeleteSubKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension\{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}\SHOWINCOMPUTER"
       
    
    Else
        MsgBox "Unknown system!", vbCritical
    End If

End Sub

Private Sub CreateCLSID(CLSID As String, Target As String, icon As String, name As String, Optional tip As String = "", Optional index As String = "")

Dim Key As RegistryKey: Set Key = Registry.ClassesRoot.CreateSubKey("CLSID\" & name)

    'Custom shell folder name : (REG_EXPAND_SZ)
    Key.SetValue "", name, REG_EXPAND_SZ
    Key.SetValue "InfoTip", tip, REG_EXPAND_SZ

    'Custom shell folder icon : (REG_EXPAND_SZ)
    Dim iconKey As RegistryKey: Set iconKey = Key.CreateSubKey("DefaultIcon")
    iconKey.SetValue "", icon, REG_EXPAND_SZ
    Set iconKey = Nothing

    'Custom shell folder required settings : "%SystemRoot%\system32\shdocvw.dll" (REG_EXPAND_SZ)
    Dim inprocServerKey As RegistryKey: Set inprocServerKey = Key.CreateSubKey("InProcServer32")
    inprocServerKey.SetValue "", "%SystemRoot%\system32\shdocvw.dll", REG_EXPAND_SZ
    inprocServerKey.SetValue "ThreadingModel", "Both", REG_SZ
    Set inprocServerKey = Nothing

    '"Folder Shortcut" ClassID
    Dim instanceKey As RegistryKey: Set instanceKey = Key.CreateSubKey("Instance")
    instanceKey.SetValue "CLSID", "{0AFACED1-E828-11D1-9187-B532F1E9575D}", REG_SZ

    'Custom shell folder real path (REG_SZ)
    Dim initPropertyBagKey As RegistryKey: Set initPropertyBagKey = instanceKey.CreateSubKey("InitPropertyBag")
    initPropertyBagKey.SetValue "Attributes", 21, REG_DWORD
    initPropertyBagKey.SetValue "Target", Target, REG_SZ
    Set initPropertyBagKey = Nothing

    'Custom shell folder attributes
    Set Key = Key.CreateSubKey("ShellFolder")
    Key.SetValue "Attributes", 671088753, REG_DWORD
    Key.SetValue "WantsFORPARSING", "", REG_SZ
    Set Key = Nothing

    'Add the custom shell folder to the Desktop
    Registry.LocalMachine.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\" & CLSID) _
        .SetValue "", "CustomShellFolder", REG_SZ

    'Add the custom shell folder to the Computer
    Registry.LocalMachine.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\" & CLSID) _
        .SetValue "", "CustomShellFolder", REG_SZ

    'Show the custom shell folder icon on the Desktop
    Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel") _
        .SetValue CLSID, 0

    'Show the custom shell folder icon in Computer
    Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\HideMyComputerIcons") _
        .SetValue CLSID, 0, REG_DWORD

    'Icon Removal message
    Registry.LocalMachine.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\" & CLSID) _
        .SetValue "Removal Message", "mydocs.dll,-900", REG_SZ


End Sub

Private Sub CreateFolderOption(CLSID As String, iconPath As String, name As String, Optional shellType As Long = 0)

    Dim Key As RegistryKey: Set Key = Registry.LocalMachine.OpenBaseKey(HKEY_LOCAL_MACHINE). _
                                                            OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension", True)
    
    If (Key Is Nothing) Then
        Set Key = Registry.LocalMachine.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShellExtension")
    End If

    Dim subKey As RegistryKey: Set subKey = Key.CreateSubKey(CLSID)

    If (shellType = 1) Then
        subKey.SetValue "RegPath", "Software\Microsoft\Windows\CurrentVersion\Explorer\HideMyComputerIcons", REG_SZ
        subKey.SetValue "Text", "Show " & name & " in Computer", REG_SZ
        subKey.SetValue "Type", "checkbox", REG_SZ
        subKey.SetValue "CheckedValue", 0, REG_DWORD
        subKey.SetValue "ValueName", CLSID, REG_SZ
        subKey.SetValue "DefaultValue", 1, REG_DWORD
        subKey.SetValue "UncheckedValue", 1, REG_DWORD
        subKey.SetValue "HKeyRoot", 2147483649#, REG_DWORD
        subKey.SetValue "Bitmap", Replace(iconPath, "-", ""), REG_SZ
    Else
        subKey.SetValue "Text", "Where to show " & name & " folder", REG_SZ
        subKey.SetValue "Type", "group", REG_SZ
        subKey.SetValue "Bitmap", Replace(iconPath, "-", ""), REG_EXPAND_SZ

        Dim showOnDesktop As RegistryKey
        Set showOnDesktop = subKey.CreateSubKey("SHOWONDESKTOP")
        
        showOnDesktop.SetValue "RegPath", "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", REG_SZ
        showOnDesktop.SetValue "Text", "Show on Desktop", REG_SZ
        showOnDesktop.SetValue "Type", "checkbox", REG_SZ
        showOnDesktop.SetValue "CheckedValue", 0, REG_DWORD
        showOnDesktop.SetValue "ValueName", CLSID, REG_SZ
        showOnDesktop.SetValue "DefaultValue", 1, REG_DWORD
        showOnDesktop.SetValue "UncheckedValue", 1, REG_DWORD
        showOnDesktop.SetValue "HKeyRoot", 2147483649#, REG_DWORD

        Dim showInComputer As RegistryKey: Set showInComputer = subKey.CreateSubKey("SHOWINCOMPUTER")
        showInComputer.SetValue "RegPath", "Software\Microsoft\Windows\CurrentVersion\Explorer\HideMyComputerIcons\NewStartPanel", REG_SZ
        showInComputer.SetValue "Text", "Show " & name & " folder in Computer", REG_SZ
        showInComputer.SetValue "Type", "checkbox", REG_SZ
        showInComputer.SetValue "CheckedValue", 0, REG_DWORD
        showInComputer.SetValue "ValueName", CLSID, REG_SZ
        showInComputer.SetValue "DefaultValue", 1, REG_DWORD
        showInComputer.SetValue "UncheckedValue", 1, REG_DWORD
        showInComputer.SetValue "HKeyRoot", 2147483649#, REG_DWORD
    End If
End Sub

Private Sub FixFolderPaths()

Dim shellFoldersKey As RegistryKey: Set shellFoldersKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", True)

    shellFoldersKey.SetValue "Personal", "%USERPROFILE%\My Documents", REG_SZ
    shellFoldersKey.SetValue "My Pictures", "%USERPROFILE%\My Pictures", REG_SZ
    shellFoldersKey.SetValue "My Music", "%USERPROFILE%\My Music", REG_SZ
    shellFoldersKey.SetValue "My Video", "%USERPROFILE%\My Video", REG_SZ
    shellFoldersKey.SetValue "Downloads", "%USERPROFILE%\Downloads", REG_SZ

Dim userShellFoldersKey As RegistryKey: Set userShellFoldersKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", True)

    userShellFoldersKey.SetValue "Personal", "%USERPROFILE%\My Documents", REG_EXPAND_SZ
    userShellFoldersKey.SetValue "My Pictures", "%USERPROFILE%\My Pictures", REG_EXPAND_SZ
    userShellFoldersKey.SetValue "My Music", "%USERPROFILE%\My Music", REG_EXPAND_SZ
    userShellFoldersKey.SetValue "My Video", "%USERPROFILE%\My Video", REG_EXPAND_SZ
    userShellFoldersKey.SetValue "Downloads", "%USERPROFILE%\Downloads", REG_EXPAND_SZ
    
Dim Key As RegistryKey: Set Key = Registry.Users.OpenSubKey(".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", True)

    Key.SetValue "Personal", "%USERPROFILE%\My Documents", REG_SZ
    Key.SetValue "My Pictures", "%USERPROFILE%\My Pictures", REG_SZ
    Key.SetValue "My Music", "%USERPROFILE%\My Music", REG_SZ
    Key.SetValue "My Video", "%USERPROFILE%\My Video", REG_SZ
    Key.SetValue "Downloads", "%USERPROFILE%\Downloads", REG_SZ
    
Set Key = Registry.Users.OpenSubKey(".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", True)

    Key.SetValue "Personal", "%USERPROFILE%\My Documents", REG_EXPAND_SZ
    Key.SetValue "My Pictures", "%USERPROFILE%\My Pictures", REG_EXPAND_SZ
    Key.SetValue "My Music", "%USERPROFILE%\My Music", REG_EXPAND_SZ
    Key.SetValue "My Video", "%USERPROFILE%\My Video", REG_EXPAND_SZ
    Key.SetValue "Downloads", "%USERPROFILE%\Downloads", REG_EXPAND_SZ
End Sub



