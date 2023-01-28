Attribute VB_Name = "URLHelper"
Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const URL_ESCAPE_SPACES_ONLY As Long = &H4000000

Private Declare Function UrlEscapeW Lib "shlwapi" ( _
    ByVal pszURL As Long, _
    ByVal pszEscaped As Long, _
    ByRef pcchEscaped As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function UrlUnescapeW Lib "shlwapi" ( _
    ByVal pszURL As Long, _
    ByVal pszUnescaped As Long, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long

Public Function URLDecode( _
    ByVal URL As String) As String
    
    Dim cchUnescaped As Long
    Dim HRESULT As Long
    
    If PlusSpace Then URL = Replace$(URL, "+", " ")
    cchUnescaped = Len(URL)
    
    URLDecode = String$(cchUnescaped, 0)
    HRESULT = UrlUnescapeW(StrPtr(URL), StrPtr(URLDecode), cchUnescaped, 0)
    
    If HRESULT = E_POINTER Then
        URLDecode = String$(cchUnescaped, 0)
        HRESULT = UrlUnescapeW(StrPtr(URL), StrPtr(URLDecode), cchUnescaped, 0)
    End If
    
    If HRESULT <> S_OK Then
        'Err.Raise Err.LastDllError, "URLUtility.URLDecode", _
                  "System error"
    End If
    
    URLDecode = Left$(URLDecode, cchUnescaped)
End Function

Public Function URLEncode( _
    ByVal URL As String) As String
    
    Dim cchEscaped As Long
    Dim HRESULT As Long
    
    cchEscaped = Len(URL) * 1.5
    URLEncode = String$(cchEscaped, 0)
    
    HRESULT = UrlEscapeW(StrPtr(URL), StrPtr(URLEncode), cchEscaped, URL_ESCAPE_SPACES_ONLY)
    
    If HRESULT = E_POINTER Then
        URLEncode = String$(cchEscaped, 0)
        HRESULT = UrlEscapeW(StrPtr(URL), StrPtr(URLEncode), cchEscaped, URL_ESCAPE_SPACES_ONLY)
    End If

    If HRESULT <> S_OK Then
        'Err.Raise Err.LastDllError, "URLUtility.URLEncode", _
                  "System error"
    End If
    
    URLEncode = Left$(URLEncode, cchEscaped)
End Function


