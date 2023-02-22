VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Data
Private m_lPtr As Long
' Data size:
Private m_lSize As Long
' Current position
Private m_lStreamPos As Long

Implements IStream

Public Sub Init(ByVal lPtr As Long, ByVal lSize As Long)
   m_lPtr = lPtr
   m_lSize = lSize
End Sub

Public Property Get IStreamOf() As VBStrm.IStream
   Set IStream = Me
End Property


Private Function IStream_Clone() As VBStrm.IStream
    '
End Function

Private Sub IStream_Commit(ByVal grfCommitFlags As VBStrm.STGC)
   '
End Sub


Private Sub IStream_CopyTo(ByVal pstm As VBStrm.IStream, ByVal cb As Currency, pcbRead As Currency, pcbWritten As Currency)
   '
End Sub


Private Sub IStream_LockRegion(ByVal libOffset As Currency, ByVal cb As Currency, ByVal dwLockType As Long)
   '
End Sub



Private Sub IStream_Read(ByVal pv As Long, ByVal cb As Long, ByVal pcbRead As Long)

   ' Check if we are within bounds
   If (m_lStreamPos + cb) > m_lSize Then
      cb = m_lSize - m_lStreamPos
   End If
    
   ' Copy the data to the pv pointer
   RtlMoveMemory ByVal pv, ByVal UnsignedAdd(m_lPtr, m_lStreamPos), cb
    
   ' Move the stream position
   m_lStreamPos = m_lStreamPos + cb
    
   ' Check if pcbRead is not NULL
   ' and copy the read byte count.
   If Not (pcbRead = 0) Then
      RtlMoveMemory ByVal pcbRead, cb, 4
   End If

End Sub

Private Sub IStream_Revert()
   '
End Sub


Private Sub IStream_Seek(ByVal dlibMove As Currency, ByVal dwOrigin As STREAM_SEEK, ByVal plibNewPosition As Long)
        
    ' Check dwOrgin and set
    ' the position according to it.
   Select Case dwOrigin
   Case STREAM_SEEK_SET
      m_lStreamPos = dlibMove
   Case STREAM_SEEK_CUR
      m_lStreamPos = m_lStreamPos + dlibMove
   Case STREAM_SEEK_END
      m_lStreamPos = m_lSize - 1 - dlibMove
   End Select
    
   ' Copy the new position if
   ' plibNewPosition is not NULL
   If plibNewPosition Then
      RtlMoveMemory ByVal plibNewPosition, CCur(m_lStreamPos), 8
   End If
    
End Sub

Private Sub IStream_SetSize(ByVal libNewSize As Currency)
   '
End Sub


Private Sub IStream_Stat(pstatstg As VBStrm.STATSTG, ByVal grfStatFlag As VBStrm.STATFLAG)
   ' Fill in the pstatstg structure:
   pstatstg.Type = STGTY_LOCKBYTES
   pstatstg.cbSize = m_lSize
   'pstatstg.mtime
   'pstatstg.ctime
   'pstatstg.atime
   pstatstg.grfMode = STGM_READ
   pstatstg.grfLocksSupported = 0
   ' not used: pstatstg.CLSID
   
   
   
   ' Say that this stream doesn't have a name
   RtlMoveMemory ByVal grfStatFlag, STATFLAG_NONAME, 4
   '
End Sub


Private Sub IStream_UnlockRegion(ByVal libOffset As Currency, ByVal cb As Currency, ByVal dwLockType As Long)
   '
End Sub


Private Sub IStream_Write(ByVal pv As Long, ByVal cb As Long, ByVal pcbWritten As Long)
   '
End Sub



