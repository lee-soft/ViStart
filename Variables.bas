Attribute VB_Name = "Variables"
Public UserVariable As New Collection

Public Function VarScan(ByVal sData As String)

Dim p1 As Integer, p2 As Integer, SVar As String, sP1 As String, sP2 As String, _
    sVarValue As String, sEnvironValue As String
    
    p1 = 1
    p2 = 1
    
    While p1 > 0
        p1 = InStr(p1, sData, "%")
        p2 = p1 + 1
        p1 = InStr(p1 + 1, sData, "%")
        
        If p1 > 0 And p2 > 0 Then
            'Found Variable
            SVar = Mid$(sData, p2, p1 - p2)

            sP1 = Mid$(sData, 1, p2 - 2)
            sP2 = Mid$(sData, Len(sP1) + Len(SVar) + 3)
            
            'Put it into data
            sVarValue = UsrVarValue(SVar, "%" & SVar & "%")
            
            'Check if variable has variable
            While IsVariable(sVarValue) = True
                sEnvironValue = Environ$(StripVarKey(sVarValue))
            
                If sEnvironValue = "" Then
                    'See if its a user variable
                    sVarValue = UsrVarValue(sVarValue, "<" & SVar & ">")
                Else
                    'global/enviroment variable
                    sVarValue = sEnvironValue
                End If
            Wend

            sData = sP1 & sVarValue & sP2
            p1 = InStr(Len(sP1) + Len(sVarValue), sData, "%")
        End If
    Wend
    
    VarScan = sData

End Function

Private Function IsVariable(ByVal sData As String) As Boolean

    If Left$(sData, 1) = "%" And Right$(sData, 1) = "%" Then
        IsVariable = True
    End If

End Function

Function StripVarKey(sValue As String)

    If Len(sValue) > 1 Then
        StripVarKey = Mid$(sValue, 2, Len(sValue) - 2)
    Else
        StripVarKey = sValue
    End If

End Function

Function UsrVarValue(sName As String, Optional Default = "") As String

    If Not ExistInCol(UserVariable, sName) Then
        UsrVarValue = Default
    
        If IsVariable(sName) = True Then
            Logger.Error "Invalid Variable", "UsrValValue", sName, Default
        End If
        Exit Function
    End If

    UsrVarValue = UserVariable(sName)
End Function

