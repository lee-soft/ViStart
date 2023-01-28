Attribute VB_Name = "WDSHelper"
Option Explicit

Private m_Connection

'Private m_busy As Boolean
'Private m_interuptQuery As Boolean
Public Function WDSAvailable() As Boolean

    On Error GoTo Handler

    Set m_Connection = CreateObject("ADODB.Connection")
    m_Connection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"

    WDSAvailable = IIf(Not m_Connection Is Nothing, True, False)
    Exit Function
Handler:
    WDSAvailable = False
    
End Function
