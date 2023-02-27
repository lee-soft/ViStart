VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TV_Type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCaption As String 'local copy
Private mvarChildren As Collection 'local copy
Private mvarNode As INode 'local copy
Private mvarDisplayLimit As Long 'local copy
Public AllowQuery As Boolean

Public Property Let DisplayLimit(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DisplayLimit = 5
    mvarDisplayLimit = vData
End Property

Public Property Get DisplayLimit() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DisplayLimit
    DisplayLimit = mvarDisplayLimit
End Property

Public Property Set Node(ByVal vData As INode)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Node = Form1
    Set mvarNode = vData
End Property


Public Property Get Node() As INode
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Node
    Set Node = mvarNode
End Property


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

Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property

Private Sub Class_Initialize()
    mvarDisplayLimit = -1
End Sub
