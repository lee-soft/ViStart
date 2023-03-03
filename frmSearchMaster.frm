VERSION 5.00
Begin VB.Form frmSearchMaster 
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form2"
   ScaleHeight     =   3360
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Text            =   "a"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmSearchMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IHookSink

Private Const NOTIFY_NEWSEARCH As Long = 1
Private Const NOTIFY_ABORTSEARCH As Long = 0

Private m_masterhWnd As Long

Private m_slaves As Collection
'Private m_theCollection As Collection

Public Event onNewItem()

Public Event onNewService(ByRef theSearchObject As CustomSearchSlave)
Public Event onUpdateCollection(ByRef theSearchObject As CustomSearchSlave)

Public Function SendAbortSignal()
    ExecuteNewQuery vbNullString
End Function

Public Function OpenSlaves()
    On Error Resume Next

Dim thisFile As Scripting.File

    If FSO.FolderExists(App.Path & "\Plugins\") Then
        For Each thisFile In FSO.GetFolder(App.Path & "\Plugins\").Files
        
            If LCase$(thisFile.Name) = "metroprovider.exe" Then
                If g_Windows8 Then
                    Shell thisFile.Path
                End If
            Else
                Shell thisFile.Path
            End If
        Next
    End If

End Function

Public Function TerminateSlaves()

Dim thisItem As CustomSearchSlave

    For Each thisItem In m_slaves
        SendAppMessage Me.hWnd, thisItem.hWnd, "BYE BYE"
    Next
End Function

Public Function ExecuteNewQuery(szNewQuery As String)
    Debug.Print "EXECUTEQUERY " & szNewQuery

Dim thisItem As CustomSearchSlave
Dim notifyCode As Long

    If szNewQuery = vbNullString Then
        notifyCode = NOTIFY_ABORTSEARCH
    Else
        notifyCode = NOTIFY_NEWSEARCH
    End If
    
    For Each thisItem In m_slaves
        'SendAppMessage CStr(thisItem), "QUERY " & URLEncode(newQuery)
        'If IsWindow(thisItem.hWnd) = APIFALSE Then
            'OpenSlaves
        'End If

        Set thisItem.Results = New Collection
        
        If notifyCode = NOTIFY_NEWSEARCH Then RaiseEvent onUpdateCollection(thisItem)
        PostMessage ByVal thisItem.hWnd, ByVal CLng(WM_USER + 1987), ByVal CLng(notifyCode), ByVal CLng(frmStartMenuBase.SearchBoxhWnd)
    Next

End Function

Private Function FlushItemsForCollection(thehWnd As String)

Dim thisItem As CustomSearchSlave

    If ExistInCol(m_slaves, "win_" & thehWnd) Then
        Set thisItem = m_slaves("win_" & thehWnd)

        While thisItem.Results.count > 0
            thisItem.Results.Remove 1
        Wend
    End If

End Function

Private Function GetServiceByName(szSearchService As String) As CustomSearchSlave

Dim thisSearchSlave As CustomSearchSlave

    For Each thisSearchSlave In m_slaves
        If UCase$(thisSearchSlave.SearchType) = UCase$(szSearchService) Then
        
            Set GetServiceByName = thisSearchSlave
            Exit Function
        End If
    Next

End Function

Private Function RecieveItemFromSlave(szCaption As String, szPath As String, szIcon As String, sourcehWnd As Long)

Dim searchSlave As CustomSearchSlave

    DoEvents

    If Not ExistInCol(m_slaves, "win_" & sourcehWnd) Then
        SendAppMessage Me.hWnd, sourcehWnd, "ERR UNREGISTERED"
        Exit Function
    End If

    Set searchSlave = m_slaves("win_" & sourcehWnd)
    AddItem searchSlave.Results, szCaption, szPath, szIcon

End Function

Private Function RegisterSearchSlave(sourcehWnd As Long, szSearchService As String)

Dim searchSlave As CustomSearchSlave

    Set searchSlave = GetServiceByName(szSearchService)

    If Not searchSlave Is Nothing Then
        If IsWindow(searchSlave.hWnd) = APIFALSE Then
            m_slaves.Remove "win_" & searchSlave.hWnd
        Else
            SendAppMessage Me.hWnd, sourcehWnd, "ERR OCCUPIED_BY_VALID_SLAVE"
            Exit Function
        End If
    End If
    
    If ExistInCol(m_slaves, "win_" & sourcehWnd) Then
        SendAppMessage Me.hWnd, sourcehWnd, "ERR THIS_HWND_HAS_SERVICE"
        Exit Function
    End If
    
    Set searchSlave = New CustomSearchSlave
    Set searchSlave.Results = New Collection
    
    searchSlave.hWnd = sourcehWnd
    searchSlave.SearchType = szSearchService

    m_slaves.Add searchSlave, "win_" & searchSlave.hWnd
    SendAppMessage Me.hWnd, sourcehWnd, "OKAY"
    
    RaiseEvent onNewService(searchSlave)
End Function

Private Function RecieveAppMessage(ByVal sourcehWnd As Long, ByVal theData As String)
    On Error GoTo Handler

Dim sP() As String

    sP = Split(theData, " ")
    
    Select Case UCase$(sP(0))
    
    Case "NEW_ORB"
        Dim proposedOrb As String: proposedOrb = URLDecode(CStr(sP(1)))

        If Not FileExists(sCon_AppDataPath & "_orbs\" & proposedOrb) Then
            Logger.Error "Attempting to apply new orb has failed due to missing file", "RecieveAppMessage"
            Exit Function
        End If
        
        frmStartOrb.Path = sCon_AppDataPath & "_orbs\" & proposedOrb
        Settings.CurrentOrb = proposedOrb
        frmControlPanel.ListOrbs
        
    Case "PIN_FILE"
        Dim proposedFile As String: proposedFile = URLDecode(CStr(sP(1)))
        Settings.Programs.TogglePin_ElseAddToPin_ByProgram CreateProgramFromPath(proposedFile)
    
    Case "NEW_THEME"
        Dim proposedSkin As String: proposedSkin = URLDecode(CStr(sP(1)))
    
        If Not FileCheck(sCon_AppDataPath & "_skins\" & proposedSkin & "\") Then
            Logger.Error "Attempting to apply new skin has failed due to missing or inaccessible files", "RecieveAppMessage"
            Exit Function
        End If
    
        frmStartMenuBase.Skin = proposedSkin
        frmControlPanel.ListSkins
    
    Case "REGISTER"
        If UBound(sP) = 1 Then
            If IsWindow(sourcehWnd) = APITRUE Then
                RegisterSearchSlave sourcehWnd, URLDecode(CStr(GetPublicString("strFiles")))
            End If
        End If
    
    Case "ITEM"
        'If m_ignoreUntilReady Then Exit Function
        
        If UBound(sP) = 3 Then
            'AddItem URLDecode(sP(1)), URLDecode(sP(2))
            RecieveItemFromSlave URLDecode(sP(1)), URLDecode(sP(2)), URLDecode(sP(3)), sourcehWnd
            Debug.Print theData
            
        ElseIf UBound(sP) = 2 Then
            RecieveItemFromSlave URLDecode(sP(1)), URLDecode(sP(2)), vbNullString, sourcehWnd
        End If
        
    Case "NEW"
        FlushItemsForCollection CStr(sourcehWnd)
        'm_ignoreUntilReady = False

    End Select
    
    Exit Function
Handler:
    If sourcehWnd <> 0 Then
        SendAppMessage Me.hWnd, sourcehWnd, "ERR UNEXPECTED_COMMAND_PARAMETER"
    End If
    
End Function

Private Function AddItem(theCollection As Collection, szCaption As String, szPath As String, Optional szIconPath As String)

Dim thisNode As New INode

    With thisNode
        .Caption = szCaption
        .IsFile = True
        .Tag = szPath
        
        If szIconPath = vbNullString Then
            'Lets the treeview component load .tag as an icon when it needs to
            .IconPosition = -3
        Else
        
            Set .Icon = IconManager.GetViIcon(szIconPath, False)
            .IconPosition = -4
        
            'Set .Icon = New ViIcon

            'If .Icon.LoadIconFromFile(szIconPath) Then
                '.IconPosition = -4
            'End If
        End If
    End With
    
    theCollection.Add thisNode

    RaiseEvent onNewItem
End Function

Private Sub Form_Load()
    Call HookWindow(Me.hWnd, Me)
    
    Set m_slaves = New Collection
    Me.Caption = MASTERID
    
    OpenSlaves
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindow Me.hWnd
    TerminateSlaves
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long

    On Error GoTo Handler
    
Dim tCDS As COPYDATASTRUCT
Dim b() As Byte

    If msg = WM_COPYDATA Then

        win.CopyMemory tCDS, ByVal lp, Len(tCDS)
        ReDim b(0 To tCDS.cbData) As Byte
        
        If tCDS.dwData = 87 Then
            win.CopyMemory b(0), ByVal tCDS.lpData, tCDS.cbData
            RecieveAppMessage wp, CStr(b)
        End If
    End If

    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       CallOldWindowProcessor(hWnd, msg, wp, lp)
       
    Exit Function
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       CallOldWindowProcessor(hWnd, msg, wp, lp)
End Function

