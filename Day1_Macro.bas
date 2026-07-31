Sub Day1_Macro()
'
' B1セルに赤の太字で「こんにちは」と表示し、1行目の高さを20に指定
'
    Range("B1").Select                     ' B1セルを選択
    ActiveCell.FormulaR1C1 = "こんにちは"   ' 文字を入力
    Range("B1").Select
    Selection.Font.Bold = True             ' 太字にする
    Selection.Font.Color = vbRed           ' 文字色を赤にする
    Rows("1:1").Select
    Selection.RowHeight = 20               ' 行の高さを20に変更
End Sub