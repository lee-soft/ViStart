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

Public Function LongToDWordByteArray(ByVal Value As Long) As Byte()
    ' Create a byte array to hold the result
    Dim result(3) As Byte
    
    ' Copy the bytes of the Long value into the byte array
    win.CopyMemory result(0), Value, 4
    
    ' Return the byte array
    LongToDWordByteArray = result
End Function

Public Function DWordByteArrayToLong(ByRef sourceByte() As Byte) As Long
    Dim myLong As Long
    
    Debug.Print sourceByte(0) & ":" & sourceByte(1) & ":" & sourceByte(2) & ":" & sourceByte(3)
    
    ' Assume myBytes contains the DWORD value in little-endian byte order
    ' (i.e., least significant byte first)
    
    ' Get a pointer to the byte array
    Dim myPtr As Long
    myPtr = StrPtr(sourceByte(0))
    
    ' Copy the bytes into a Long variable
    win.CopyMemory myLong, myPtr, 4
    
    'ByteSwap (result)

    DWordByteArrayToLong = myLong
End Function

Public Function ConvertBytesToLong(bytes() As Byte) As Long
    Dim result As Long
    
    ' Copy the bytes into the result variable
    win.CopyMemory result, bytes(0), 4
    
    ' Return the result
    ConvertBytesToLong = result
End Function

