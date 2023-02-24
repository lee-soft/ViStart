VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIconManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal iconIndex As Long, ByVal Flags As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long

Private m_LargeIcons As Long
Private m_SmallIcons As Long

Private m_iconLib As Collection

Private m_extensions As Collection
Private m_links As Collection

Private m_folderIconSmall As ViIcon
Private m_folderIconLarge As ViIcon

Private Sub Class_Initialize()
    Set m_iconLib = New Collection
    Set m_extensions = New Collection
    Set m_links = New Collection
    
    m_LargeIcons = ImageList_Create(32, 32, ILC_COLOR32 Or ILC_MASK, 1, 16)
    m_SmallIcons = ImageList_Create(16, 16, ILC_COLOR32 Or ILC_MASK, 1, 16)
    
    Set m_folderIconSmall = New ViIcon: m_folderIconSmall.Handle = SHGetIcon(Environ$("windir"), False)
    Set m_folderIconLarge = New ViIcon: m_folderIconLarge.Handle = SHGetIcon(Environ$("windir"), True)
    
End Sub

'Gets a ViIcon from a path, ensuring it's not loaded twice
Public Function GetViIcon(ByVal szPath As String, Optional bLarge As Boolean) As ViIcon

    If szPath = vbNullString Then
        Exit Function
    End If

Dim szExtension As String
Dim returnObj As ViIcon

Dim potentialNewIcon As Long

Dim iconIndex As Long
Dim szIconPath As String

Dim szKey As String
Dim szExtKey As String

    szPath = PathRemoveBlackSlash(szPath)
    szPath = UCase$(szPath)
    szExtension = PathFindExtension(szPath)
    
    szKey = IIf(bLarge, "_large:", "_small:") & szPath
    szExtKey = IIf(bLarge, "_large:", "_small:") & szExtension
    
    If szExtension <> vbNullString Then
        If Not szExtension = "LNK" And Not szExtension = "EXE" And Not szExtension = "ICO" And Not szExtension = "URL" Then
            If Not ExistInCol(m_extensions, szExtKey) Then
                Set returnObj = New ViIcon
                returnObj.Handle = SHGetIcon(szPath, bLarge)

                m_extensions.Add returnObj, szExtKey
            Else
                Set returnObj = m_extensions(szExtKey)
            End If
        Else
            If Not ExistInCol(m_links, szKey) Then
                Set returnObj = New ViIcon
            
                If GetIconLocationAndIndex(szPath, iconIndex, szIconPath) = False Then
                    returnObj.Handle = SHGetIcon(szPath, bLarge)
                Else
                    potentialNewIcon = GetIcon(szIconPath, iconIndex, bLarge)
  
                    If potentialNewIcon = 0 Then
                        potentialNewIcon = SHGetIcon(szPath, bLarge)
                    End If
                    
                    returnObj.Handle = potentialNewIcon
                End If

                m_links.Add returnObj, szKey
            Else
                Set returnObj = m_links(szKey)
            End If
        End If
    Else
        If Not bLarge Then
            Set returnObj = m_folderIconSmall
        Else
            Set returnObj = m_folderIconLarge
        End If
    End If
    
    If returnObj Is Nothing Then
        Set returnObj = New ViIcon
    End If
    
    returnObj.IconPath = szPath
    Set GetViIcon = returnObj
End Function

Private Function AddIcon(ByVal hIcon As Long, ByVal bLarge As Boolean) As Long
Dim res As Long

    If (hIcon <> 0) Then
        res = ImageList_AddIcon(IIf(bLarge, m_LargeIcons, m_SmallIcons), hIcon)
        DestroyIcon hIcon

        AddIcon = ImageList_GetIcon(IIf(bLarge, m_LargeIcons, m_SmallIcons), res, ILD_TRANSPARENT)
    End If
End Function

Private Function GetIcon(ByVal szIconLocation As String, Optional iconIndex As Long = -1, Optional bLarge As Boolean = False) As Long

Dim iconArray(0) As Long

    szIconLocation = Wow64Wrapper(szIconLocation)

    If bLarge Then
        If ExtractIconExW(StrPtr(szIconLocation), iconIndex, iconArray(0), 0, 1) = 0 Then
            Exit Function
        End If
    Else
        If ExtractIconExW(StrPtr(szIconLocation), iconIndex, 0, iconArray(0), 1) = 0 Then
            Exit Function
        End If
    End If
    
    GetIcon = AddIcon(iconArray(0), bLarge)
End Function

Private Function SHGetIcon(ByVal szPath As String, Optional bLarge As Boolean) As Long

Dim hIcon As Long
Dim res As Long

    szPath = Wow64Wrapper(szPath)
    
    hIcon = SHExtractIcon(szPath, IIf(bLarge, ICON_SIZE.LARGE_ICON, ICON_SIZE.SMALL_ICON))
    
    SHGetIcon = AddIcon(hIcon, bLarge)
End Function

Private Sub Class_Terminate()
    ImageList_Destroy m_LargeIcons
    ImageList_Destroy m_SmallIcons
End Sub
