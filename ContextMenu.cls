VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_vistaMenu As frmVistaMenu
Attribute m_vistaMenu.VB_VarHelpID = -1

Public Id As String

Public Function PopupMenu(ownerForm As Form) As Form
    Set PopupMenu = m_vistaMenu
    m_vistaMenu.Resurrect True, ownerForm
End Function

Public Function PopulateViaXML(contextMenuObj As IXMLDOMElement) As Boolean

Dim thisItem As IXMLDOMElement
Dim thisObject As Object
    
    Set m_vistaMenu = New frmVistaMenu
    
    If Not AttributeExists(contextMenuObj, "id") Then
        PopulateViaXML = False
        Exit Function
    End If
    
    Id = contextMenuObj.Attributes.getNamedItem("id").Text

    For Each thisObject In contextMenuObj.childNodes
        If TypeName(thisObject) = "IXMLDOMElement" Then
            Set thisItem = thisObject
        
            If thisItem.nodeName = "item" Then
                m_vistaMenu.AddItem thisItem.Attributes.getNamedItem("caption").Text, thisItem.Attributes.getNamedItem("program").Text
            End If
        End If
    Next

End Function

Private Sub m_vistaMenu_onClick(theItemTag As String)
    m_vistaMenu.Hide

    If Is64bit Then
    
        Dim win64Token As New Win64FSToken
        win64Token.EnableFS
        'ExplorerRun theFilePath
        If AppLauncherHelper.ShellEx(theItemTag) = APIFALSE Then Exit Sub
        
        Set win64Token = Nothing
    Else
        If AppLauncherHelper.ShellEx(theItemTag) = APIFALSE Then Exit Sub
    End If
End Sub

Private Sub m_vistaMenu_onInActive()
    m_vistaMenu.Hide
End Sub
