Attribute VB_Name = "DWordHelper"
Option Explicit

Public Declare Function CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (dest As Any, _
                                      Src As Any, _
                                      ByVal cb As Long) As Long

Public Function LoWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(LoWord, LongIn, 2)
End Function

Public Function HiWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Public Function MakeLong(ByVal HiWord As Integer, ByVal LoWord As Integer) As Long
   Call CopyMemory(MakeLong, LoWord, 2)
   Call CopyMemory(ByVal (VarPtr(MakeLong) + 2), HiWord, 2)
End Function

Public Function LongToDWordByteArray(ByVal Value As Long) As Byte()
    ' Create a byte array to hold the result
    Dim result(3) As Byte
    
    ' Copy the bytes of the Long value into the byte array
    CopyMemory result(0), Value, 4
    
    ' Return the byte array
    LongToDWordByteArray = result
End Function

Public Function DWordByteArrayToLong(ByRef sourceByte() As Byte) As Long
    Dim myLong As Long
    
    ' Assume myBytes contains the DWORD value in little-endian byte order
    ' (i.e., least significant byte first)
    
    ' Get a pointer to the byte array
    Dim myPtr As Long
    myPtr = StrPtr(sourceByte(0))
    
    ' Copy the bytes into a Long variable
    CopyMemory myLong, myPtr, 4
    
    'ByteSwap (result)

    DWordByteArrayToLong = myLong
End Function

Public Function ULongTovbLong(Value As Double) As Long

Const OFFSET_4 = 4294967296#
Const MAXINT_4 = &H7FFFFFFF

    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    If Value <= MAXINT_4 Then
        ULongTovbLong = Value
    Else
        ULongTovbLong = Value - OFFSET_4
    End If
End Function

