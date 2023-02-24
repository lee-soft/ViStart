VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPPrivateFC"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ==================================================================================
' cPrivateFontCollection.cls
' Copyright ? 2003 Genghis Khan(GuangJian Guo)
' Website: http://www.itkhan.com
'
' from GdiplusFontCollection.h:Microsoft Platform SDK February 2003
'
' The PrivateFontCollection is a collection class for fonts. This class keeps a
' collection of fonts specifically for an application. The fonts in the collection
' can include installed fonts as well as fonts that have not been installed on the system.
' ==================================================================================

Option Explicit

Private m_uLastStatus As GpStatus

Private m_clsFontColl As GDIPFontCollection

Public Property Get FontCollection() As GDIPFontCollection
    Set FontCollection = m_clsFontColl
End Property

Public Function AddFontFile(ByVal FileName As String) As GpStatus
    SetStatus GdipPrivateAddFontFile(m_clsFontColl.NativeFontCollection, FileName)
    AddFontFile = m_uLastStatus
    
End Function

Private Function SetStatus(ByVal status As GpStatus) As GpStatus
    m_uLastStatus = SetStatusHelper(status)
End Function

Public Function GetLastStatus() As GpStatus

    GetLastStatus = m_uLastStatus
End Function

Public Property Get NativeFontCollection() As Long
    NativeFontCollection = m_clsFontColl.NativeFontCollection
End Property

Private Sub Class_Initialize()

    Dim lCol As Long

    SetStatus GdipNewPrivateFontCollection(lCol)
    Set m_clsFontColl = New GDIPFontCollection
    m_clsFontColl.fInitialize (lCol)
End Sub

Private Sub Class_Terminate()
    Call GdipDeletePrivateFontCollection(m_clsFontColl.NativeFontCollection)
    Set m_clsFontColl = Nothing
End Sub
