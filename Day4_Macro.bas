Sub ExtractRows(wsSrc As Worksheet, wsDst As Worksheet, col As Long, keyword As String)
    '
    ' 指定列から keyword を含む行だけ抽出して別シートにコピーする
    '
    
    Dim lastRow As Long
    Dim dstRow As Long
    Dim i As Long

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, col).End(xlUp).Row

    dstRow = 1
    
    For i = 1 To lastRow
        If InStr(wsSrc.Cells(i, col).Value, keyword) > 0 Then
            wsSrc.Rows(i).Copy wsDst.Rows(dstRow)
            dstRow = dstRow + 1
        End If
    Next i
End Sub

Function SheetExists(sheetName As String) As Boolean
    '
    ' 指定したシート名が存在するかどうかを判定する
    '
    
    On Error Resume Next
    SheetExists = Not Worksheets(sheetName) Is Nothing   ' 存在すれば True
    On Error GoTo 0
End Function

Sub Day4_Macro()

    ' 元データのシート名と抽出先シート名
    Dim srcName As String: srcName = "Sheet1"
    Dim dstName As String: dstName = "Sheet2"

    ' シート存在チェック（どちらか無ければ終了）
    If Not SheetExists(srcName) Or Not SheetExists(dstName) Then
        MsgBox "指定されたシートが存在しません。処理を中止します。"
        Exit Sub
    End If

    ' シートが存在するので抽出処理を実行
    ExtractRows Worksheets(srcName), Worksheets(dstName), 1, "りんご"

    MsgBox "抽出が完了しました。"
End Sub