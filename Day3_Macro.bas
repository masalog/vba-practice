Sub Day3_Macro()
'
' Sheet1をSheet2にコピー
'

Application.DisplayAlerts = False
Worksheets("Sheet2").Delete
Application.DisplayAlerts = True

Worksheets("Sheet1").Copy After:=Worksheets("Sheet1")
ActiveSheet.Name = "Sheet2"

End Sub