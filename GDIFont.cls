VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Const DEFAULT_CHARSET As Long = 1

'local variable(s) to hold property value(s)
Private mvarHandle As Long 'local copy
Private mvarFontFace As String 'local copy
Private mvarFontWeight As Long 'local copy
Private mvarItalic As BOOL 'local copy

Public Property Get Handle() As Long
Attribute Handle.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Handle
    Handle = mvarHandle
End Property

Public Property Get FontWeight() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontWeight
    FontWeight = mvarFontWeight
End Property

Public Property Get FontFace() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontFace
    FontFace = mvarFontFace
End Property

Public Property Get Italic() As BOOL
    Italic = mvarItalic
End Property

Public Function Constructor(Optional ByVal FontFace As String = "Tahoma", Optional ByVal FontWeight As Long = 15, Optional ByVal Italic As BOOL = APIFALSE)
    If mvarHandle <> 0 Then
        DeleteObject mvarHandle
    End If
    
    mvarFontFace = FontFace
    mvarFontWeight = FontWeight
    mvarItalic = Italic
    
    mvarHandle = CreateFont(15, 0, 0, 0, mvarFontWeight, Italic, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, mvarFontFace)
End Function

Private Sub Class_Initialize()
    Constructor "Tahoma", 0
End Sub

Private Sub Class_Terminate()
    If mvarHandle <> 0 Then
        DeleteObject mvarHandle
    End If
End Sub
