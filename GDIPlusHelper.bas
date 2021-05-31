Attribute VB_Name = "GDIPlusHelper"
Public Function CopyRectL(ByRef srcRect As gdiplus.RECTL, ByRef dstRect As gdiplus.RECTL)

    dstRect.Left = srcRect.Left
    dstRect.Top = srcRect.Top
    dstRect.Height = srcRect.Height
    dstRect.Width = srcRect.Width

End Function

Public Function CreateRectF(Left As Single, Top As Single, Height As Single, Width As Single) As gdiplus.RECTF

Dim newRectF As gdiplus.RECTF

    With newRectF
        .Left = Left
        .Top = Top
        .Height = Height
        .Width = Width
    End With
    
    CreateRectF = newRectF
End Function

