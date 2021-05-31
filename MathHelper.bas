Attribute VB_Name = "MathHelper"
Option Explicit

Public Function FindIndex(mousePosition As Long, dividerAmount As Long) As Double

    If mousePosition < dividerAmount Then
        FindIndex = 0
    Else
        FindIndex = Floor(mousePosition / dividerAmount)
    End If

End Function

Public Function Ceiling(dblValue As Double) As Double
'Returns the smallest integral value that is greater than or equal to the specified double-precision floating-point number.

On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
    Ceiling = CDbl(Left$(CStr(dblValue), myDec)) + 1
Else
    Ceiling = dblValue
End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox Err.Description, vbInformation, "Round Up"
End Function

Public Function Floor(dblValue As Double) As Double
'Returns the largest integer less than or equal to the specified number.

On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
    Floor = CDbl(Left$(CStr(dblValue), myDec))
Else
    Floor = dblValue
End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox Err.Description, vbInformation, "Round Down"
End Function

Public Function RoundIt(ByVal lngSrcNumber As Integer, ByVal lngByNumber As Integer)

    'Round(12, 5)
   
Dim lngModResult As Long
    lngModResult = (lngSrcNumber Mod lngByNumber)
   
    If lngModResult >= lngByNumber Then
        RoundIt = CLng(SymArith(lngSrcNumber / lngByNumber, 0) * lngByNumber + 1)
    Else
        RoundIt = CLng(SymArith(lngSrcNumber / lngByNumber, 0) * lngByNumber)
    End If

End Function

Private Function SymArith(ByVal X As Double, _
  Optional ByVal DecimalPlaces As Double = 1) As Double

    SymArith = Fix(X * (10 ^ DecimalPlaces) + 0.5 * Sgn(X)) _
        / (10 ^ DecimalPlaces)
End Function

