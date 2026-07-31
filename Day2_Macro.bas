Sub Day2_Macro()
'
' A1〜A10に1〜10を自動入力、偶数→青、奇数→緑に色分け
'
    Dim i As Long
    Dim c As Range

    ' A1〜A10に1〜10を入力
    For i = 1 To 10
        Range("A" & i).Value = i
    Next i

    ' 選択範囲の値が偶数→青、奇数→緑に色分け
    For Each c In Selection
        If IsNumeric(c.Value) Then
            If c.Value Mod 2 = 0 Then
                c.Interior.Color = vbBlue
            Else
                c.Interior.Color = vbGreen
            End If
        End If
    Next c

End Sub
