VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private Const m_cNodeSpace As Long = 19

Private Const m_cIconSize As Long = 16


Private mRoot As INode
Attribute mRoot.VB_VarHelpID = -1
Private mcAllNodes As Collection
Private mcSpacePositions As Collection

Private sDebugPadding As String
Private bAbort As Boolean
Private bCalculateSize As Boolean

Private mLngLevel As Long
Private mLngBiggestHeight As Long

Private m_NodeCount As Long

Public Event onNode(targetNode As INode)

Public Event onDownLevel()
Public Event onUpLevel()

Private mvarID As String 'local copy
Public Property Let Id(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Function ColapseAll()
    
Dim thisNode As INode

    For Each thisNode In mcAllNodes
    
        If thisNode.Expanded = True Then
            thisNode.Expanded = False
        End If
    Next
    
End Function

Public Function ColapseChildren(ByRef srcNode As INode)
    
Dim thisNode As INode

    For Each thisNode In srcNode.Children
    
        If thisNode.Expanded = True Then
            thisNode.Expanded = False
        End If
        
        If thisNode.Children.count > 0 Then
            ColapseChildren thisNode
        End If
    Next
    
End Function

Public Function ClearNodes()
    Set mcAllNodes = New Collection
End Function

Public Property Get AllNodes() As Collection
    Set AllNodes = mcAllNodes
End Property

Public Property Get RootNode() As INode
    Set RootNode = mRoot
End Property

Public Function Abort()
    bAbort = True
End Function

Public Function Notify_Expanded()
    bCalculateSize = True
End Function

Public Property Get Size() As RECT

Dim newSize As RECT

    If bCalculateSize Then
        bCalculateSize = False
        
        mLngBiggestHeight = 0
        GetRootSize mRoot
    End If
    
    newSize.Bottom = mLngBiggestHeight
    Size = newSize

End Property

Private Function GetRootSize(srcNode As INode)

Dim Node As INode
Dim SubNode As INode

    If srcNode.Expanded Then
        For Each Node In srcNode.Children
            If Not Node.visibleInSearchOnly Then
                mLngLevel = mLngLevel + m_cIconSize
                GetSize Node
                mLngLevel = mLngLevel - m_cIconSize
            End If
        Next
    End If

End Function

Private Function GetSize(srcNode As INode)

Dim Node As INode
Dim SubNode As INode

    'If srcNode.IconPosition = -1 Then
        'mLngBiggestHeight = mLngBiggestHeight + m_cNodeSpace + M_SEPARATOR_GAP
        'Debug.Print "mLngBiggestHeight:: " & mLngBiggestHeight
    'Else
        mLngBiggestHeight = mLngBiggestHeight + m_cNodeSpace
    'End If
        
    m_NodeCount = m_NodeCount + 1

    If srcNode.Expanded Then
        For Each Node In srcNode.Children
            GetSize Node
        Next
    End If

End Function

Public Function irriterateNode(ByRef srcNode As INode)

Dim Node As INode
Dim SubNode As INode

    bAbort = False
    
    'If srcNode.Expanded Then
        For Each Node In srcNode.Children
            If bAbort Then
                Exit For
            End If

            If Not Node.visibleInSearchOnly Then
                debugPrint Node
            End If
        Next
    'End If

End Function

'replaced with above routine because it was causing items to appear above the bitmaptext
'
'Public Function irriterateNode(ByRef srcNode As INode)

'Dim Node As INode
'Dim SubNode As INode

'    bAbort = False
    
'    For Each Node In srcNode.Children
'        If Node.IsFile Then
    
'            If bAbort Then
'                Exit For
'            End If

'            debugPrint Node
'        End If
'    Next
    
'    For Each Node In srcNode.Children
'        If Not Node.IsFile Then
    
'            If bAbort Then
'                Exit For
'            End If

'            debugPrint Node
'        End If
'   Next

'End Function

Private Function debugPrint(Optional ByRef srcNode As INode)

Dim Node As INode
Dim SubNode As INode

    RaiseEvent onNode(srcNode)
    
    If srcNode.Expanded Then
        For Each Node In srcNode.Children
            If bAbort Then
                Exit For
            End If
            
            RaiseEvent onUpLevel
            debugPrint Node
            RaiseEvent onDownLevel
        Next
    End If

End Function

Public Function createNode(Optional ByRef srcNode As INode) As INode
    
Dim new_Node As INode
    
    Set new_Node = New INode
    
    If (srcNode Is Nothing) Then
        Set srcNode = mRoot
    End If
    
    srcNode.Children.Add new_Node
    mcAllNodes.Add new_Node
    
    srcNode.Root = Me
    Set createNode = new_Node
    
End Function

Public Function ShowAll(ByRef srcCollection As Collection, Optional resultLimiter As Long = -1, Optional ByRef exceedsLimits As Boolean) As RECT
    If srcCollection Is Nothing Then
        Exit Function
    End If
    
    bAbort = False
    
Dim nNode As INode
Dim lngNewWidth As Long
Dim LngBiggestHeight As Long
Dim LngBiggestWidth As Long
Dim resultCount As Long
    
    
    For Each nNode In srcCollection
        lngNewWidth = nNode.Width + m_cNodeSpace
        LngBiggestHeight = LngBiggestHeight + m_cNodeSpace
    
        If lngNewWidth > LngBiggestWidth Then
            LngBiggestWidth = lngNewWidth
        End If
        
        RaiseEvent onNode(nNode)
        
        resultCount = resultCount + 1
        If resultCount = resultLimiter Or bAbort Then
            exceedsLimits = True
            Exit For
        End If
    Next
    
    ShowAll.Right = LngBiggestWidth
    ShowAll.Bottom = LngBiggestHeight
End Function

Public Function QueryCollection(ByVal strKeyword As String, ByRef srcCollection As Collection, Optional intResultLimit As Long = -1, Optional bExeedsResult As Boolean, Optional ByRef lngSpacesLeftOver As Long) As RECT
    If srcCollection Is Nothing Then
        Exit Function
    End If
    
    'Bug, Query Node wont investigate its children
    'only the nodes inside of itself

Dim lngNodeIndex As Long
Dim lngKeyWordIndex As Long
Dim lngByteSpaceIndex As Long

Dim strItemName As String

Dim bSearchSpacePositions() As Long
Dim bKeyWordSpacePositions() As Long

Dim lngLenItemName As Long
Dim lngLenKeyWords() As Long

Dim aStrKeyWords() As String

Dim bShow() As Boolean
Dim bShowFinal As Boolean

Dim nNode As INode
Dim childNode As INode

Dim lngNewWidth As Long
Dim newSize As RECT

Dim LngBiggestWidth As Long
Dim LngBiggestHeight As Long

Dim lngResultCount As Long
Dim exeKeyWord As String

    exeKeyWord = Replace(UCase$(strKeyword), " ", "")

    aStrKeyWords = Split(UCase$(strKeyword), " ")
    ReDim lngLenKeyWords(UBound(aStrKeyWords))
    
    For lngKeyWordIndex = 0 To UBound(aStrKeyWords)
        lngLenKeyWords(lngKeyWordIndex) = Len(aStrKeyWords(lngKeyWordIndex))
    Next
    
    LngBiggestHeight = 0
    LngBiggestWidth = 0
    lngResultCount = 0
    
    For Each nNode In srcCollection
        If lngResultCount = intResultLimit Then
            bExeedsResult = True
            Exit For
        End If
    
        'For each Node with no children
        If nNode.Children.count = 0 Then
            bSearchSpacePositions = nNode.SpacePositions
            
            If IsValidArray(bSearchSpacePositions) Then
                strItemName = nNode.SearchIdentifier
                
                lngLenItemName = Len(strItemName)
                
                ReDim bShow(UBound(aStrKeyWords))
                bShowFinal = True
                
                'For Each First Letter
                For lngByteSpaceIndex = 0 To UBound(bSearchSpacePositions)
        
                    For lngKeyWordIndex = 0 To UBound(aStrKeyWords)
                        If Mid$(strItemName, bSearchSpacePositions(lngByteSpaceIndex), lngLenKeyWords(lngKeyWordIndex)) = aStrKeyWords(lngKeyWordIndex) Then
                            bShow(lngKeyWordIndex) = True
                        End If
                    Next
                Next
                
                
                For lngKeyWordIndex = 0 To UBound(aStrKeyWords)
                    If bShow(lngKeyWordIndex) = False Then
                        bShowFinal = False
                        Exit For
                    End If
                Next
                
                If bShowFinal = False Then
                    If Mid$(nNode.EXEName, 1, Len(exeKeyWord)) = exeKeyWord Then
                        bShowFinal = True
                    End If
                End If
                
                If bShowFinal Then
    
                    lngNewWidth = nNode.Width + m_cNodeSpace
                    LngBiggestHeight = LngBiggestHeight + m_cNodeSpace
            
                    If lngNewWidth > LngBiggestWidth Then
                        LngBiggestWidth = lngNewWidth
                    End If
                
                    lngResultCount = lngResultCount + 1
                    RaiseEvent onNode(nNode)
                End If
            Else
            
                'Debug.Print "Suspect: " & nNode.Caption
            End If
        End If
    Next
    
    lngSpacesLeftOver = intResultLimit - lngResultCount
    
    newSize.Bottom = LngBiggestHeight
    newSize.Right = LngBiggestWidth

    QueryCollection = newSize
    
End Function

Private Sub Class_Initialize()
    
    Set mRoot = New INode
    Set mcAllNodes = New Collection
    
    mRoot.Root = Me
    mRoot.Expanded = True
    
    mRoot.Caption = "Root"
    
End Sub
