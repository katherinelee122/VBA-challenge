Sub StockCounter()
Dim Ticker As String
Dim TotalVolume As Variant
Dim YearlyChange As Variant
Dim PercentChange As Variant
Dim OpenPrice As Variant
Dim ClosingPrice As Variant
Dim GreatestIncrease As Variant
Dim GreatestDecrease As Variant
Dim GreatestTotalVolume As Variant
For Each ws In Worksheets
SummaryTableRow = 2
TotalVolume = 0
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
OpeningPrice = ws.Cells(2, 3)
For i = 2 To LastRow
   If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    Ticker = ws.Cells(i, 1)
    TotalVolume = TotalVolume + ws.Cells(i, 7)
    ClosingPrice = ws.Cells(i, 6)
    YearlyChange = ClosingPrice - OpeningPrice
   If OpeningPrice = 0 Then
       PercentChange = 0
   Else
       PercentChange = YearlyChange / OpeningPrice * 100
   End If
ws.Range("i" & SummaryTableRow).Value = Ticker
ws.Range("j" & SummaryTableRow).Value = YearlyChange
ws.Range("k" & SummaryTableRow).Value = PercentChange
ws.Range("l" & SummaryTableRow).Value = TotalVolume
TotalVolume = 0
OpeningPrice = ws.Cells(i + 1, 3)
If ws.Range("J" & SummaryTableRow).Value >= 0 Then
   ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
Else
   ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
   End If
SummaryTableRow = SummaryTableRow + 1
Else
TotalVolume = TotalVolume + ws.Cells(i, 7)
End If

If ws.Cells(i, 9).Value > GreatestIncrease Then
GreatestIncrease = ws.Cells(i, 9).Value
ws.Cells(2, 17).Value = ws.Cells(i, 9).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
End If

If ws.Cells(i, 9).Value < GreatestDecrease Then
GreatestDecrease = ws.Cells(i, 9).Value
ws.Cells(3, 17).Value = ws.Cells(i, 9).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
End If

If ws.Cells(i, 12).Value > GreatestTotalVolume Then
GreatestTotalVolume = ws.Cells(i, 12).Value
ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
End If

Next i
   ws.Range("I1") = "Ticker"
   ws.Range("P1") = "Ticker"
   ws.Range("J1") = "Yearly Change"
   ws.Range("K1") = "Percent Change"
   ws.Range("L1") = "Total Volume"
   ws.Range("Q1") = "Value"
   ws.Range("O2") = "Greatest Percent Increase"
   ws.Range("O3") = "Greatest Percent Decrease"
   ws.Range("O4") = "Greatest Total Volume"
   


    Next ws
End Sub
