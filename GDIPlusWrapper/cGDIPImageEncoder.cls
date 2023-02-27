VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPImageCodec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_iClsId As clsid
Private m_iName As String
Private m_iDllName As String
Private m_iExt As String
Private m_iFlags As Long
Private m_iDescription As String
Private m_iFormatId As clsid
Private m_iMimeType As String
Private m_iVersion As Long

Public Property Get CodecCLSID() As clsid
   CodecCLSID = m_iClsId
End Property
Public Property Get CodecName() As String
   CodecName = m_iName
End Property
Public Property Get DllName() As String
   DllName = m_iDllName
End Property
Public Property Get FilenameExtension() As String
   FilenameExtension = m_iExt
End Property
Public Property Get Flags() As Long
   Flags = m_iFlags
End Property
Public Property Get FormatDescription() As String
   FormatDescription = m_iDescription
End Property
Public Property Get FormatID() As clsid
   FormatID = m_iFormatId
End Property
Public Property Get MimeType() As String
   MimeType = m_iMimeType
End Property
Public Property Get Version() As Long
   Version = m_iVersion
End Property

Friend Sub fInit( _
      iClsId As clsid, _
      iName As String, _
      iDllName As String, _
      iExt As String, _
      iFlags As Long, _
      iDescription As String, _
      iFormatId As clsid, _
      iMimeType As String, _
      iVersion As Long _
   )
   LSet m_iClsId = iClsId
   m_iName = iName
   m_iDllName = iDllName
   m_iExt = iExt
   m_iFlags = iFlags
   m_iDescription = iDescription
   LSet m_iFormatId = m_iFormatId
   m_iMimeType = iMimeType
   m_iVersion = iVersion
End Sub
