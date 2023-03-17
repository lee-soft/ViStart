Attribute VB_Name = "MemoryToPictureHelper"
Option Explicit

Private Enum CBoolean   ' enum members are Long data types
  CFalse = 0
  CTrue = 1
End Enum
 
Private Const S_OK = 0    ' indicates successful HRESULT
 
'WINOLEAPI CreateStreamOnHGlobal(
'    HGLOBAL hGlobal,            // Memory handle for the stream object
'    BOOL fDeleteOnRelease,  // Whether to free memory when the object is released
'    LPSTREAM * ppstm           // Indirect pointer to the new stream object
');
Private Declare Function CreateStreamOnHGlobal Lib "ole32" _
                              (ByVal hGlobal As Long, _
                              ByVal fDeleteOnRelease As CBoolean, _
                              ppstm As Any) As Long
 
'STDAPI OleLoadPicture(
'    IStream * pStream,  // Pointer to the stream that contains picture's data
'    LONG lSize,            // Number of bytes read from the stream
'    BOOL fRunmode,   // The opposite of the initial value of the picture's property
'    REFIID riid,             // Reference to the identifier of the interface describing the type
'                                   // of interface pointer to return
'    VOID ppvObj          // Indirect pointer to the object, not AddRef'd!!
');
Private Declare Function OleLoadPicture Lib "olepro32" _
                              (pStream As Any, _
                              ByVal lSize As Long, _
                              ByVal fRunmode As CBoolean, _
                              riid As GUID, _
                              ppvObj As Any) As Long
 
Private Type GUID    ' 16 bytes (128 bits)
  dwData1 As Long      ' 4 bytes
  wData2 As Integer     ' 2 bytes
  wData3 As Integer     ' 2 bytes
  abData4(7) As Byte   ' 8 bytes, zero based
End Type
 
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long
 
Private Const sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
 
Private Const GMEM_MOVEABLE = &H2
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
 
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
 
' ====================================================================
 
 
Private Type OPENFILENAME  '  ofn
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
 
Public Function PictureFromBits(abPic() As Byte) As IPicture  ' not a StdPicture!!

Dim nLow As Long
Dim cbMem  As Long
Dim hMem  As Long
Dim lpMem  As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown '  IStream
Dim ipic As IPicture

    ' Get the size of the picture's bits
    
    nLow = LBound(abPic)
    
    cbMem = (UBound(abPic) - nLow) + 1
    
    ' Allocate a global memory object
    hMem = GlobalAlloc(GMEM_MOVEABLE, cbMem)
    
    If hMem Then
        
        ' Lock the memory object and get a pointer to it.
        lpMem = GlobalLock(hMem)
        
        If lpMem Then
            
            ' Copy the picture bits to the memory pointer and unlock the handle.
            MoveMemory ByVal lpMem, abPic(nLow), cbMem
            Call GlobalUnlock(hMem)
            
            ' Create an ISteam from the pictures bits (we can explicitly free hMem
            ' below, but we'll have the call do it...)
            If (CreateStreamOnHGlobal(hMem, CTrue, istm) = S_OK) Then
                If (CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture) = S_OK) Then
                    
                    ' Create an IPicture from the IStream (the docs say the call does not
                    ' AddRef its last param, but it looks like the reference counts are correct..)
                    Call OleLoadPicture(ByVal ObjPtr(istm), cbMem, CFalse, _
                    IID_IPicture, PictureFromBits)
                    
                End If
            End If
        End If
        
        'Call GlobalFree(hMem)
    End If
    
End Function

