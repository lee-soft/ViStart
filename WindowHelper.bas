Attribute VB_Name = "WindowHelper"
Option Explicit

Public Function GetWindowClassNameByHwnd(ByVal hWnd As Long) As String

Dim lReturn As Long
Dim mvarCaption As String

    mvarCaption = Space$(256)
    lReturn = GetClassName(hWnd, mvarCaption, Len(mvarCaption))
    
    If lReturn Then
        GetWindowClassNameByHwnd = Left$(mvarCaption, lReturn)
    End If

End Function

Public Function GetWindowNameByHwnd(ByVal hWnd As Long) As String

Dim lReturn As Long
Dim mvarCaption As String

    mvarCaption = Space$(256)
    lReturn = GetWindowText(hWnd, mvarCaption, Len(mvarCaption))
    
    If lReturn Then
        GetWindowNameByHwnd = Left$(mvarCaption, lReturn)
    End If

End Function
