Attribute VB_Name = "Module1"
Sub MultipleYearStockData():
For Each ws In Worksheets
Dim WorksheetName As String
Dim i As Long
Dim j As Long
Dim TickCount As Long
Dim LastRow1 As Long
Dim LastRow2 As Long
Dim PercentChange As Double
Dim Increase As Double
Dim Decrease As Double
Dim Volumn As Double
WorksheetName = ws.Name
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
TickCount = 2
j = 2
LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow1
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Cells(TickCount, 9).Value = Cells(i, 1).Value
Cells(TickCount, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value
If Cells(TickCount, 10).Value < 0 Then
Cells(TickCount, 10).Interior.Color = RGB(200, 0, 0)
Else
Cells(TickCount, 10).Interior.Color = RGB(0, 200, 0)
End If
If Cells(j, 3).Value <> 0 Then
PercentChange = ((Cells(i, 6).Value - Cells(j, 3).Value) / Cells(j, 3).Value)
Cells(TickCount, 11).Value = Format(PercentChange, "Percent")
Else
Cells(TickCount, 11).Value = Format(0, "Percent")
End If
Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(Cells(j, 7), Cells(i, 7)))
TickCount = TickCount + 1
j = i + 1
End If
Next i
LastRow2 = Cells(Rows.Count, 9).End(xlUp).Row
Volumn = Cells(2, 12).Value
Increase = Cells(2, 11).Value
Decrease = Cells(2, 11).Value
For i = 2 To LastRow2
If Cells(i, 12).Value > Volumn Then
Volumn = Cells(i, 12).Value
Cells(4, 16).Value = Cells(i, 9).Value
Else
Volumn = Volumn
End If
If Cells(i, 11).Value > Increase Then
Increase = Cells(i, 11).Value
Cells(2, 16).Value = Cells(i, 9).Value
Else
Increase = Increase
End If
If Cells(i, 11).Value < Decrease Then
Decrease = Cells(i, 11).Value
Cells(3, 16).Value = Cells(i, 9).Value
Else
Decrease = Decrease
End If
Cells(2, 17).Value = Format(Increase, "Percent")
Cells(3, 17).Value = Format(Decrease, "Percent")
Cells(4, 17).Value = Format(Volumn, "Scientific")
Next i
Worksheets(WorksheetName).Columns("I:Q").AutoFit
Next ws
End Sub




