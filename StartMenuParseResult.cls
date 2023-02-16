VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StartMenuParseResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarStartMenuPath As String 'local copy
Private mvarUserFramePath As String 'local copy
Private mvarBottomButtonsArrowPath As String 'local copy
Private mvarBottomButtonsShutdownPath As String 'local copy
Private mvarAllProgramsPath As String 'local copy
Private mvarButtonPath As String 'local copy
Private mvarProgramsArrowPath As String 'local copy
Private mvarErrorCode As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarStartMenuMaskPath As String 'local copy
Private mvarStartMenuExpandedPath As String 'local copy
Public Property Let StartMenuExpandedPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartMenuExpandedPath = 5
    mvarStartMenuExpandedPath = vData
End Property


Public Property Get StartMenuExpandedPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartMenuExpandedPath
    StartMenuExpandedPath = mvarStartMenuExpandedPath
End Property



Public Property Let StartMenuMaskPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartMenuMaskPath = 5
    mvarStartMenuMaskPath = vData
End Property


Public Property Get StartMenuMaskPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartMenuMaskPath
    StartMenuMaskPath = mvarStartMenuMaskPath
End Property



Public Property Let ErrorCode(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ErrorCode = 5
    mvarErrorCode = vData
End Property


Public Property Get ErrorCode() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorCode
    ErrorCode = mvarErrorCode
End Property



Public Property Let ProgramsArrowPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProgramsArrowPath = 5
    mvarProgramsArrowPath = vData
End Property


Public Property Get ProgramsArrowPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProgramsArrowPath
    ProgramsArrowPath = mvarProgramsArrowPath
End Property



Public Property Let ButtonPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ButtonPath = 5
    mvarButtonPath = vData
End Property


Public Property Get ButtonPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ButtonPath
    ButtonPath = mvarButtonPath
End Property



Public Property Let AllProgramsPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AllProgramsPath = 5
    mvarAllProgramsPath = vData
End Property


Public Property Get AllProgramsPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AllProgramsPath
    AllProgramsPath = mvarAllProgramsPath
End Property



Public Property Let BottomButtonsShutdownPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BottomButtonsShutdownPath = 5
    mvarBottomButtonsShutdownPath = vData
End Property


Public Property Get BottomButtonsShutdownPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BottomButtonsShutdownPath
    BottomButtonsShutdownPath = mvarBottomButtonsShutdownPath
End Property



Public Property Let BottomButtonsArrowPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BottomButtonsArrowPath = 5
    mvarBottomButtonsArrowPath = vData
End Property


Public Property Get BottomButtonsArrowPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BottomButtonsArrowPath
    BottomButtonsArrowPath = mvarBottomButtonsArrowPath
End Property



Public Property Let UserFramePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserFramePath = 5
    mvarUserFramePath = vData
End Property


Public Property Get UserFramePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserFramePath
    UserFramePath = mvarUserFramePath
End Property



Public Property Let StartMenuPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartMenuPath = 5
    mvarStartMenuPath = vData
End Property


Public Property Get StartMenuPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartMenuPath
    StartMenuPath = mvarStartMenuPath
End Property



