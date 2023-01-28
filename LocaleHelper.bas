Attribute VB_Name = "LocaleHelper"
Private Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Const LOCALE_SENGLANGUAGE As Long = &H1001

Public Function GetUserLocale() As String
Dim strReturn As String

    On Error GoTo English:

    strReturn = GetUserLocaleInfo(GetSystemDefaultLCID(), LOCALE_SENGLANGUAGE)
    GetUserLocale = strReturn
    
    Exit Function
English:
    GetUserLocale = "English"
End Function

Private Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function
