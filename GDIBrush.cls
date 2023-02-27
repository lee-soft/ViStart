VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Colour"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private m_Brush As Long

Public Property Let Colour(theColour As Long)
    Constructor theColour
End Property

Public Property Get Value()
    Value = m_Brush
End Property

Public Function Constructor(brushColour As Long)
    Class_Terminate
    m_Brush = CreateSolidBrush(brushColour)
End Function

Private Sub Class_Terminate()
    If m_Brush <> 0 Then
        DeleteObject ByVal m_Brush
        m_Brush = 0
    End If
End Sub
