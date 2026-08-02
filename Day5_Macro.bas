Function CopyRowsAboveThreshold(wsSrc As Worksheet, wsDst As Worksheet, _
                                lastRow As Long, threshold As Long, _
                                checkCol As Long, copyStartCol As Long, _
                                copyColCount As Long)
    
    '指定した列の値が閾値以上なら、その行の指定列範囲をコピーする
    
    Dim i As Long
    Dim dstRow As Long
    
    '見出しは1行目にコピー済み
    dstRow = 2

    For i = 2 To lastRow
        If wsSrc.Cells(i, checkCol).Value >= threshold Then

            'Value を使う（高速化）
            wsDst.Cells(dstRow, copyStartCol).Resize(1, copyColCount).Value = _
                wsSrc.Cells(i, copyStartCol).Resize(1, copyColCount).Value

            dstRow = dstRow + 1
        End If
    Next i

End Function


Function SortByColumnAscending(ws As Worksheet, col As Long)
    
    '指定した列を昇順で並べ替える（1行目は見出しとして除外）
    
    Dim rng As Range
    Dim colCount As Long

    'データが見出しのみならソート不要
    If ws.UsedRange.Rows.Count <= 1 Then
        Exit Function
    End If

    colCount = ws.UsedRange.Columns.Count

    '★ 2行目から最終行の範囲を自動取得
    Set rng = ws.Range("A2").Resize(ws.UsedRange.Rows.Count - 1, colCount)

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Cells(2, col), Order:=xlAscending

    ws.Sort.SetRange rng
    ws.Sort.Header = xlNo
    ws.Sort.Apply

End Function


Sub Day5_Macro()
    
    '閾値以上の売上データを別シートに転記する
    

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastRow As Long
    Dim threshold As Variant

    Set wsSrc = Sheets("Sheet1")
    Set wsDst = Sheets("Sheet2")
    
    '閾値を入力
    threshold = InputBox("閾値となる売上を入力してください", "閾値入力")

    '数字チェック
    If Not IsNumeric(threshold) Then
        MsgBox "数字を入力してください。", vbExclamation
        Exit Sub
    End If

    '出力先を初期化
    wsDst.Cells.Clear

    lastRow = wsSrc.Cells(Rows.Count, 1).End(xlUp).Row
    
    '★ 見出しコピー
    wsSrc.Rows(1).Copy wsDst.Rows(1)

    '★ 閾値以上の行をコピー（C列＝3列目を判定）
    CopyRowsAboveThreshold wsSrc, wsDst, lastRow, CLng(threshold), 3, 1, 3

    '★ A列を昇順で並べ替え
    SortByColumnAscending wsDst, 1

    MsgBox "売上整理が完了しました！"

End Sub