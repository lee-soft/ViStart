Attribute VB_Name = "DWordHelper"
Option Explicit

Public Function LOWORD(ByVal LongIn As Long) As Integer
   Call win.CopyMemory(LOWORD, LongIn, 2)
End Function

Public Function HiWord(ByVal LongIn As Long) As Integer
   Call win.CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Public Function MakeLong(ByVal HiWord As Integer, ByVal LOWORD As Integer) As Long
   Call win.CopyMemory(MakeLong, LOWORD, 2)
   Call win.CopyMemory(ByVal (VarPtr(MakeLong) + 2), HiWord, 2)
End Function

Public Function LongToDWordByteArray(ByVal value As Long) As Byte()
    ' Create a byte array to hold the result
    Dim result(3) As Byte
    
    ' Copy the bytes of the Long value into the byte array
    win.CopyMemory result(0), value, 4
    
    ' Return the byte array
    LongToDWordByteArray = result
End Function
