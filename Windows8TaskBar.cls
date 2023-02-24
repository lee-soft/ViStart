VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Windows8TaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_originalRect As gdiplus.RECTL
Private m_resizedRect As gdiplus.RECTL

Private m_lastKnownEdge As AbeBarEnum

'Windows 7 StartOrb Gap
Const STARTBAR_GAP As Long = 45

Private Function CheckEdge()

Dim currentTaskBarEdge As AbeBarEnum

    currentTaskBarEdge = GetTaskBarEdge()
    
    If m_lastKnownEdge <> currentTaskBarEdge Then
        m_lastKnownEdge = currentTaskBarEdge
        
        GetRespectiveRect g_hwndMSTask, m_originalRect
        ReMakeResizedTaskBar
    End If

End Function

Function CheckTaskbarSize()

Dim currentTaskBarRect As gdiplus.RECTL

    CheckEdge
    GetRespectiveRect g_hwndMSTask, currentTaskBarRect
    
    If m_lastKnownEdge = abe_bottom Or m_lastKnownEdge = ABE_TOP Then
    
        If currentTaskBarRect.Left <> m_resizedRect.Left Then
             'MoveWindow g_hwndReBarWindow32, m_resizedRect.Left, m_resizedRect.Top, m_resizedRect.Width, m_resizedRect.Height, True
             MoveWindow g_hwndMSTask, m_resizedRect.Left, m_resizedRect.Top, m_resizedRect.Width, currentTaskBarRect.Height, True
             'MoveWindow &H200C2, m_resizedRect.Left, m_resizedRect.Top, m_resizedRect.Width, m_resizedRect.Height, True
    
        End If
        
    Else
        
        If currentTaskBarRect.Top <> m_resizedRect.Top Then
             'MoveWindow g_hwndReBarWindow32, m_resizedRect.Left, m_resizedRect.Top, m_resizedRect.Width, m_resizedRect.Height, True
             MoveWindow g_hwndMSTask, m_resizedRect.Left, m_resizedRect.Top, currentTaskBarRect.Width, m_resizedRect.Height, True
             'MoveWindow &H200C2, m_resizedRect.Left, m_resizedRect.Top, m_resizedRect.Width, m_resizedRect.Height, True
    
        End If
        
    End If

End Function

Private Function ReMakeResizedTaskBar()
 
    m_resizedRect = m_originalRect

    If m_lastKnownEdge = abe_bottom Or m_lastKnownEdge = ABE_TOP Then
        
        If m_resizedRect.Left <> STARTBAR_GAP Then
            m_resizedRect.Width = m_resizedRect.Width - STARTBAR_GAP
            m_resizedRect.Left = STARTBAR_GAP
        End If
    
    Else
        
        If m_resizedRect.Top <> STARTBAR_GAP Then
            m_resizedRect.Height = m_resizedRect.Height - STARTBAR_GAP
            m_resizedRect.Top = STARTBAR_GAP
        End If
        
    End If

End Function

Private Function GetRespectiveRect(ByVal targetHwnd As Long, ByRef sourceRect As gdiplus.RECTL)

Dim originalRect As RECT
Dim parentRect As RECT
Dim parentHwnd As Long
    
    If targetHwnd = 0 Then Exit Function
    parentHwnd = getParent(targetHwnd)
    If parentHwnd = 0 Then Exit Function
    
    GetWindowRect targetHwnd, originalRect
    GetWindowRect parentHwnd, parentRect
    
    sourceRect.Top = originalRect.Top - parentRect.Top
    sourceRect.Left = originalRect.Left - parentRect.Left
    sourceRect.Width = originalRect.Right - originalRect.Left
    sourceRect.Height = originalRect.Bottom - originalRect.Top

End Function

Private Sub Class_Initialize()
    m_lastKnownEdge = -1

    If g_hwndReBarWindow32 = 0 Then Exit Sub
    
    CheckEdge
    CheckTaskbarSize
End Sub

Private Sub Class_Terminate()
    MoveWindow g_hwndMSTask, m_originalRect.Left, m_originalRect.Top, m_originalRect.Width, m_originalRect.Height, True
End Sub
