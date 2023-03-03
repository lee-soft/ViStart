VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShellLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Win32 API declarations
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

' API Constants
Private Const MAX_PATH As Long = 2024

Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINNOACTIVE = 7

' Member variables
Private m_ShellLink As ShellLinkA       ' An explorer IShellLinkA(Win 95/Win NT) instance
'Private m_shellLink2 As ivbshell

Private m_PersistFile As IPersistFile   ' An explorer IPersistFile instance
Private m_FileName As String            ' Filename
Private m_AutoQuote As Boolean          ' Automatically add quotes around targets that contain spaces?
Private m_Overwrite As Boolean          ' Determines whether create overwrites existing files
Private m_Dirty As Boolean              ' Flag indicating object needs to be saved

' Property defaults
Private Const defOverwrite As Boolean = False
Private Const defAutoQuote As Boolean = True

' *************************************************
'  Initialization / Termination
' *************************************************
Private Sub Class_Initialize()
   ' Set property defaults.
   m_Overwrite = defOverwrite
   m_AutoQuote = defAutoQuote
   ' Create new IShellLink interface
   Set m_ShellLink = New ShellLinkA
   ' Implement ShellLinkA's IPersistFile interface
   Set m_PersistFile = m_ShellLink
End Sub

Private Sub Class_Terminate()
   ' Clean up.
   Set m_PersistFile = Nothing
   Set m_ShellLink = Nothing
End Sub

' *************************************************
'  Public Properties (read-write)
' *************************************************
Public Property Get Arguments() As String
   Dim Buffer As String
   With m_ShellLink
      Buffer = Space$(MAX_PATH)
      Call .GetArguments(Buffer, Len(Buffer))
      Arguments = TrimNull(Buffer)
   End With
End Property

Public Property Let Arguments(ByVal args As String)
   With m_ShellLink
      If args <> Me.Arguments Then
         .SetArguments args
         m_Dirty = True
      End If
   End With
End Property

Public Property Get AutoQuote() As Boolean
   AutoQuote = m_AutoQuote
End Property

Public Property Let AutoQuote(ByVal NewVal As Boolean)
   m_AutoQuote = NewVal
End Property

Public Property Get Description() As String
   Dim Buffer As String
   With m_ShellLink
      Buffer = Space$(MAX_PATH)
      Call .GetDescription(Buffer, Len(Buffer))
      Description = TrimNull(Buffer)
   End With
End Property

Public Property Let Description(ByVal Desc As String)
   With m_ShellLink
      If Desc <> Me.Description Then
         .SetDescription Desc
         m_Dirty = True
      End If
   End With
End Property

Public Property Get Hotkey() As Long
   ' *** pwHotkey ***
   ' Address of the hot key. The virtual key code is in the
   ' low-order byte, and the modifier flags are in the
   ' high-order byte. The modifier flags can be a combination
   ' of the following values:
   '    HOTKEYF_ALT      ALT key
   '    HOTKEYF_CONTROL  CTRL key
   '    HOTKEYF_EXT      Extended key
   '    HOTKEYF_SHIFT    SHIFT key
   With m_ShellLink
      Call .GetHotkey(Hotkey)
   End With
End Property

Public Property Let Hotkey(ByVal NewKey As Long)
   With m_ShellLink
      If NewKey <> Me.Hotkey Then
         .SetHotkey NewKey
         m_Dirty = True
      End If
   End With
End Property

Public Property Get IconLocation(Optional index As Long = 0) As String
   Dim Buffer As String
   ' Index can't be ByVal, because it's modified by
   ' GetIconLocation call, and returned to caller.
   With m_ShellLink
      ' Request file information from ShellLink object
      Buffer = Space$(MAX_PATH)
      Call .GetIconLocation(Buffer, Len(Buffer), index)
      ' Expand any evars and trim nulls from buffer.
      Call ExpandEnvironmentStrings(Buffer, Buffer, Len(Buffer))
      IconLocation = TrimNull(Buffer)
   End With
End Property

Public Property Let IconLocation(Optional index As Long = 0, ByVal IconFile As String)
   ' This file must exist, or we'll error out down the line.
   With m_ShellLink
      If IsFile(IconFile) Then
         .SetIconLocation IconFile, index
         m_Dirty = True
      End If
   End With
End Property

Public Property Get Overwrite() As Boolean
   Overwrite = m_Overwrite
End Property

Public Property Let Overwrite(ByVal NewVal As Boolean)
   m_Overwrite = NewVal
End Property

Public Property Get Target() As String
   Dim wfd As IShellLinkA.WIN32_FIND_DATA
   Dim Buffer As String
   Const SLGP_RAWPATH As Long = 4
   
   With m_ShellLink
      ' Request file information from ShellLink object
      Buffer = Space$(MAX_PATH)
      Call .GetPath(Buffer, Len(Buffer), wfd, SLGP_RAWPATH)
      ' Expand any evars and trim nulls from buffer.
      Call ExpandEnvironmentStrings(Buffer, Buffer, Len(Buffer))
      ' Remove quotes if wanted.
      If m_AutoQuote Then
         Target = QuotesRemove(TrimNull(Buffer))
      Else
         Target = TrimNull(Buffer)
      End If
   End With
End Property

Public Property Get Exists(Optional FileSpec As String) As Boolean
   ' Default to LNK file, but accept any spec.
   If FileSpec = "" Then FileSpec = m_FileName
   Exists = IsFile(FileSpec)
End Property

Public Function QuotesAdd(ByVal TargetFile As String)
   If Len(TargetFile) > 0 Then
      TargetFile = Trim$(TargetFile)
      ' Only proceed if we have embedded space(s)
      If InStr(TargetFile, " ") Then
         If Left$(TargetFile, 1) <> """" Then
            TargetFile = """" & TargetFile
         End If
         If Right$(TargetFile, 1) <> """" Then
            TargetFile = TargetFile & """"
         End If
      End If
   End If
   QuotesAdd = TargetFile
End Function

Public Function QuotesRemove(ByVal TargetFile As String)
   TargetFile = Trim$(TargetFile)
   If Len(TargetFile) > 2 Then
      If Left$(TargetFile, 1) = """" Then
         If Right$(TargetFile, 1) = """" Then
            TargetFile = Mid$(TargetFile, 2, Len(TargetFile) - 2)
         End If
      End If
   End If
   QuotesRemove = TargetFile
End Function

Public Function Load(ByVal LinkFile As String) As Boolean
   On Error GoTo Handler
   
   Const STGM_DIRECT As Long = 0&
   
   ' Start with a fresh object.
   Call RecreateObject
   
   ' Make sure we have an accessible file.
   If IsFile(LinkFile) Then
      
      ' Load Shortcut file...(must do this UNICODE hack!)
      m_PersistFile.Load StrConv(LinkFile, vbUnicode), STGM_DIRECT
      
      ' Return success code?
      If IsFile(LinkFile) Then
         m_FileName = LinkFile
         Load = True
      End If
   End If
   
   Exit Function
Handler:
   Logger Err.Description, "Load"
End Function

' ***********************************************
'  Private Methods
' ***********************************************
Private Function IsFile(ByVal SpecIn As String) As Boolean
   Dim attr As Long
   ' Guard against bad SpecIn by ignoring errors.
   On Error Resume Next
   ' Get attribute of SpecIn.
   attr = GetAttr(SpecIn)
   If Err = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      If (attr And vbDirectory) = vbDirectory Then
         IsFile = False
      Else
         IsFile = True
      End If
   End If
End Function

Private Function PersistObject(ByVal LnkFile As String, ByVal Remember As Boolean) As Boolean
   ' TargetFile must exist, or we'll
   ' error out when Resolve is called.
   If Not IsFile(Me.Target) Then
      Exit Function
   End If
   
   ' Attempt to create as requested, making sure
   ' to do the stupid Unicode hack with filename.
   'On Error Resume Next
   m_ShellLink.Resolve 0&, SLR_UPDATE
   m_PersistFile.Save StrConv(LnkFile, vbUnicode), Remember
   
   ' Check if file exists.
   If IsFile(LnkFile) Then
      ' Only persist this filename if it's also
      ' remains as the current object, and not
      ' as a clone.
      If Remember Then
         m_FileName = LnkFile
         m_Dirty = False
      End If
      ' Return success code.
      PersistObject = True
   End If
End Function

Private Sub RecreateObject()
   ' Destory and recreate new instances of
   ' ShellLinkA and IPersistFile interfaces.
   Set m_PersistFile = Nothing
   Set m_ShellLink = New ShellLinkA
   Set m_PersistFile = m_ShellLink
   ' Reset persisted properties.
   m_Dirty = False
   m_FileName = ""
End Sub

Private Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left$(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim$(StrIn)
   End Select
End Function
