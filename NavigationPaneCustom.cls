VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavigationPaneCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_parent As NavigationPaneOption
Private m_contextMenu As frmVistaMenu

Public OpenAsMenu As Boolean

Public Property Get OptionMenu() As frmVistaMenu
    Set OptionMenu = m_contextMenu
End Property

Public Function DumpContextMenu(ByRef sourceDoc As DOMDocument, ByRef parentElement As IXMLDOMElement)

Dim thisItemIndex As Long
Dim thisItemCaption As String
Dim thisItemStyle As String
Dim thisItemExec As String

Dim XML_custom As IXMLDOMElement

    For thisItemIndex = 1 To m_contextMenu.CountItems
    
    
        thisItemCaption = m_contextMenu.GetItemCaption(thisItemIndex)
        thisItemExec = m_contextMenu.GetItemExec(thisItemIndex)
        thisItemStyle = ""
        
        If thisItemCaption = "" Then
            Set XML_custom = sourceDoc.createElement("separator")
        Else
            Set XML_custom = sourceDoc.createElement("option")
            
            If m_contextMenu.IsItemBold(thisItemIndex) Then
                thisItemStyle = "bold"
            End If
            
            XML_custom.setAttribute "caption", thisItemCaption
            XML_custom.setAttribute "style", thisItemStyle
            XML_custom.setAttribute "exec", thisItemExec
        End If
        
        parentElement.appendChild XML_custom
    Next
End Function

Public Function Populate(ByRef customElement As IXMLDOMElement)
    If Not m_contextMenu Is Nothing Then Unload m_contextMenu
    
    Set m_contextMenu = New frmVistaMenu

Dim thisObjectXML As IXMLDOMElement
Dim thisObject As Object

Dim thisCaption As String
Dim thisIsBold As Boolean
Dim thisExec As String

    For Each thisObjectXML In customElement.childNodes
        If TypeName(thisObject) = "IXMLDOMElement" Then
            Set thisObjectXML = thisObject
    
            If thisObjectXML.tagName = "option" Then
                thisIsBold = False
                thisExec = "{undefined}"
                thisCaption = "{undefined}"
    
                If Not IsNull(thisObjectXML.getAttribute("style")) Then thisIsBold = IIf(LCase$(thisObjectXML.getAttribute("style")) = "bold", True, False)
                If Not IsNull(thisObjectXML.getAttribute("caption")) Then thisCaption = thisObjectXML.getAttribute("caption")
                If Not IsNull(thisObjectXML.getAttribute("exec")) Then thisExec = thisObjectXML.getAttribute("exec")
                
            ElseIf thisObjectXML.tagName = "separator" Then
            
                thisCaption = ""
                thisExec = ""
            
            End If
            
            m_contextMenu.AddItem thisCaption, thisExec, thisIsBold
        End If
    Next
    
End Function

Public Property Get Shell() As String
    Shell = m_parent.Shell
End Property

Public Property Let Shell(newShell As String)
    m_parent.Shell = newShell
End Property

Public Property Let Visible(newVisible)
    m_parent.Visible = newVisible
End Property

Public Property Get Visible()
    Visible = m_parent.Visible
End Property

Public Property Get Position()
    Position = m_parent.Position
End Property

Public Property Let Position(newPosition)
    m_parent.Position = newPosition
End Property

Public Property Get Rollover()
    Rollover = m_parent.Rollover
End Property

Public Property Let Rollover(newRollover)
    m_parent.Rollover = newRollover
End Property

Public Property Get Caption()
    Caption = m_parent.Caption
End Property

Public Property Let Caption(newCaption)
    m_parent.Caption = newCaption
End Property

Private Sub Class_Initialize()
    Set m_parent = New NavigationPaneOption
    Set m_contextMenu = New frmVistaMenu
End Sub

