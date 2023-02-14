Attribute VB_Name = "DWordHelper"
Option Explicit

Public Function LOWORD(ByVal LongIn As Long) As Integer
   Call CopyMemory(LOWORD, LongIn, 2)
End Function

Public Function HiWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Public Function MakeLong(ByVal HiWord As Integer, ByVal LOWORD As Integer) As Long
   Call CopyMemory(MakeLong, LOWORD, 2)
   Call CopyMemory(ByVal (VarPtr(MakeLong) + 2), HiWord, 2)
End Function
