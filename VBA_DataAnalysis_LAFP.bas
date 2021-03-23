Attribute VB_Name = "Module1"
Sub wallstreet()

Dim ticker, increaseticker, decreaseticker, totalticker As String
Dim stockopen, stockclose, summarytablerow, yearlychange, percentchange, increasepercent, decreasepercent As Double
Dim stockvolume, increasevolume As LongLong

lastRow = Cells(Rows.Count, 1).End(xlUp).Row
SummaryRow = 2
stockopen = Cells(2, 3).Value

Range("I:L").ColumnWidth = 18
Range("K:K").NumberFormat = "0.00%"
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To lastRow
    If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
        ticker = Cells(i, 1).Value
        stockclose = Cells(i, 6).Value
        yearlychange = stockclose - stockopen
        If (stockopen = 0) Then
            percentchange = 0
        Else
            percentchange = yearlychange / stockopen
        End If
        stockvolume = stockvolume + Cells(i, 7)
        Cells(SummaryRow, 9).Value = ticker
        Cells(SummaryRow, 10).Value = yearlychange
        If (Cells(SummaryRow, 10).Value > 0) Then
            Cells(SummaryRow, 10).Interior.ColorIndex = 4
        ElseIf (Cells(SummaryRow, 10).Value < 0) Then
            Cells(SummaryRow, 10).Interior.ColorIndex = 3
        End If
        Cells(SummaryRow, 11).Value = percentchange
        Cells(SummaryRow, 12).Value = stockvolume
        SummaryRow = SummaryRow + 1
        stockopen = Cells(i + 1, 3).Value
        stockvolume = 0
    Else
        stockvolume = stockvolume + Cells(i, 7)
    End If
Next i

lastRow2 = Cells(Rows.Count, 11).End(xlUp).Row
increaseticker = Cells(2, 9).Value
increasepercent = Cells(2, 11).Value
decreaseticker = Cells(2, 9).Value
decreasepercent = Cells(2, 11).Value
totalticker = Cells(2, 9).Value
stockvolume = Cells(2, 12).Value

Range("O:O").ColumnWidth = 21
Range("Q2:Q3").NumberFormat = "0.00%"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For j = 2 To lastRow2
    If (Cells(j, 11).Value > increasepercent) Then
        increaseticker = Cells(j, 9).Value
        increasepercent = Cells(j, 11).Value
    End If
    If (Cells(j, 11).Value < decreasepercent) Then
        decreaseticker = Cells(j, 9).Value
        decreasepercent = Cells(j, 11).Value
    End If
    If (Cells(j, 12).Value > increasevolume) Then
        totalticker = Cells(j, 9).Value
        increasevolume = Cells(j, 12).Value
    End If
Next j

Cells(2, 16).Value = increaseticker
Cells(3, 16).Value = decreaseticker
Cells(4, 16).Value = totalticker
Cells(2, 17).Value = increasepercent
Cells(3, 17).Value = decreasepercent
Cells(4, 17).Value = increasevolume

End Sub
