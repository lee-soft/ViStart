VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsTreeview"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarIdentifier As String 'local copy
Private mvarTag As String 'local copy
Private mvarChildren As Collection 'local copy
Private mvarExpanded As Boolean 'local copy
Private mvarIconPosition As Long 'local copy
Private mvarclsTreeview As clsTreeview
Private mvarSpacePositions() As Long 'local copy
Private mvarWidth As Long 'local copy
Private mRoot As clsTreeview
Private mvarLeft As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarSearchIdentifier As String 'local copy

Public IsFile As Boolean
Public EXEName As String
Public visibleInSearchOnly As Boolean
'Public Icon As IPicture
Public Icon As ViIcon

Public Property Let SearchIdentifier(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchIdentifier = 5
    mvarSearchIdentifier = vData
    GetSpaces vData
End Property


Public Property Get SearchIdentifier() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchIdentifier
    SearchIdentifier = mvarSearchIdentifier
End Property

Public Property Let Left(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Lefts = 5
    mvarLeft = vData
End Property

Public Property Get Left() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Lefts
    Left = mvarLeft
End Property

Public Property Let Root(ByRef new_Root As clsTreeview)
    Set mRoot = new_Root
End Property

Public Property Let Width(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Width = 5
    mvarWidth = vData
End Property

Public Property Get Width() As Long
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Width = 5
    Width = mvarWidth
End Property




Public Property Get SpacePositions() As Long()
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpacePositions
    SpacePositions = mvarSpacePositions()
End Property

Public Sub GetSpaces(ByRef sKeyWord As String)

Dim bItemName() As Byte

Dim lngByteIndex As Long
Dim lngSpaceCount As Long

    'Use my 3D array (Vb has no support for 3d array ?)
    bItemName = LCase$(sKeyWord)

    ReDim Preserve mvarSpacePositions(0)
    mvarSpacePositions(lngSpaceCount) = 1
    lngSpaceCount = lngSpaceCount + 1

    For lngByteIndex = 0 To UBound(bItemName) - 2
        If bItemName(lngByteIndex) = 32 Then
            ReDim Preserve mvarSpacePositions(lngSpaceCount)
            mvarSpacePositions(lngSpaceCount) = (lngByteIndex / 2) + 2
            
            lngSpaceCount = lngSpaceCount + 1
        End If
    Next

End Sub

Public Property Get clsTreeview() As clsTreeview
    If mvarclsTreeview Is Nothing Then
        Set mvarclsTreeview = New clsTreeview
    End If


    Set clsTreeview = mvarclsTreeview
End Property

Public Property Set clsTreeview(vData As clsTreeview)
    Set mvarclsTreeview = vData
End Property

Private Sub Class_Terminate()
    Set mvarclsTreeview = Nothing
End Sub

Public Property Let IconPosition(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PositionY = 5
    mvarIconPosition = vData
End Property

Public Property Get IconPosition() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PositionY
    IconPosition = mvarIconPosition
End Property

Public Property Let Expanded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Expanded = 5
    mvarExpanded = vData
    
    'Expanded Changed, Notify the Treeview Root
    mRoot.Notify_Expanded
End Property

Public Property Get Expanded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Expanded
    Expanded = mvarExpanded
End Property

Public Sub copyNode(ByRef nTarget As INode)

Dim new_Node As INode
    
    Set new_Node = New INode
    new_Node.Root = mRoot
    
    Me.Children.Add new_Node
    mRoot.AllNodes.Add new_Node
    
    With new_Node
        .Caption = nTarget.Caption
        .Tag = nTarget.Tag
        .IconPosition = nTarget.IconPosition
        .Width = nTarget.Width
        .IsFile = nTarget.IsFile
        .EXEName = nTarget.EXEName
        
        If Not nTarget.Icon Is Nothing Then
            Set .Icon = nTarget.Icon
        End If
    End With

End Sub

Public Function createNode(Optional strCaption As String, Optional strSearchIdentifier As String, Optional strPath As String, Optional intIconPos As Long, Optional strID As String, Optional isNodeAFile As Boolean) As INode
    
Dim new_Node As INode
    
    If ExistInCol(Me.Children, strID) Then
        Set createNode = Me.Children(strID)
        Exit Function
    End If
    
    
    Set new_Node = New INode
    new_Node.Root = mRoot
    new_Node.IsFile = isNodeAFile

    Me.Children.Add new_Node, strID
    mRoot.AllNodes.Add new_Node
    
    If Not IsEmpty(strCaption) Then
        With new_Node
            .Caption = strCaption
            .SearchIdentifier = strSearchIdentifier
            
            .Tag = strPath
            .IconPosition = intIconPos

        End With
        
    End If
    
    Set createNode = new_Node


    
End Function

Public Property Set Children(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Children = Form1
    Set mvarChildren = vData
End Property


Public Property Get Children() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Children
    Set Children = mvarChildren
End Property



Public Property Let Tag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property


Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5

    mvarIdentifier = vData
End Property


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarIdentifier
End Property



Private Sub Class_Initialize()
    Set mvarChildren = New Collection
End Sub
