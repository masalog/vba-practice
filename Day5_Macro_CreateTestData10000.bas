Sub Day5_Macro_CreateTestData10000()

    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = Sheets("Sheet1")
    ws.Cells.Clear
    
    '見出し
    ws.Range("A1").Value = "商品名"
    ws.Range("B1").Value = "日付"
    ws.Range("C1").Value = "売上"
    
    '10000行のデータを作成（2026年の日付を循環）
    For i = 2 To 10001
        ws.Cells(i, 1).Value = "商品" & i
        
        '2026年1月1日から365日を循環
        ws.Cells(i, 2).Value = DateSerial(2026, 1, 1) + ((i - 2) Mod 365)
        
        '売上は1000〜20000のランダム値
        ws.Cells(i, 3).Value = Int((20000 - 1000 + 1) * Rnd + 1000)
    Next i
    
    MsgBox "2026年の10000行テストデータを作成しました！"

End Sub
