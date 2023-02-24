VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPMatrix"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : GDIPMatrix

'    Description: A Matrix object represents a 3 ×3 matrix that, in turn,
'    represents an affine transformation. A Matrix object stores only six of
'    the 9 numbers in a 3 ×3 matrix because all 3 ×3 matrices that represent
'    affine transformations have the same third column (0, 0, 1).
'
'    Based on Microsoft's GDI+ class   :
'    https://msdn.microsoft.com/en-us/library/windows/desktop/ms534475(v=vs.85).aspx
'--------------------------------------------------------------------------------


Private m_matrix As Long

Friend Property Get Value()
    Value = m_matrix
End Property

'--------------------------------------------------------------------------------
' Procedure  :       Constructor2
' Description:       Creates and initializes a Matrix::Matrix object based on
'                    six numbers that define an affine transformation.
' Parameters :       M11 (Single)
'                    M12 (Single)
'                    M21 (Single)
'                    M22 (Single)
'                    Dx (Single)
'                    Dy (Single)
'--------------------------------------------------------------------------------
Public Function Constructor2(M11 As Single, _
                             M12 As Single, _
                             M21 As Single, _
                             M22 As Single, _
                             Dx As Single, _
                             Dy As Single)
    Class_Terminate
    SetStatus GdipCreateMatrix2(M11, M12, M21, M22, Dx, Dy, m_matrix)
End Function

Private Function SetStatus(ByVal status As GpStatus) As GpStatus
    m_lastResult = SetStatusHelper(status)
End Function

Private Sub Class_Initialize()
    SetStatus GdipCreateMatrix(m_matrix)
End Sub

Private Sub Class_Terminate()

    If m_matrix <> 0 Then
        SetStatus GdipDeleteMatrix(m_matrix)
        m_matrix = 0
    End If

End Sub
