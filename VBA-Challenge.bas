Attribute VB_Name = "Module1"
Sub stockticker()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
'Start with a clean slate
ws.Range("I:Q").Clear
finalrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
'insert column headers Ticker, Yearly Change, Percent Change, Total Stock Volume
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
'insert greatest % increase, decrease, total volume
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
'create variables for Ticker, Yearly Change, Percent Change, Total Stock Volume, start price, end price
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim StartPrice As Double
Dim EndPrice As Double
Dim CurrentMax As Double
Dim CurrentMin As Double
Dim MaxVol As Double
printcount = 2
TotalStockVolume = 0
For i = 2 To finalrow
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        StartPrice = ws.Cells(i, 3).Value
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ws.Cells(printcount, 9).Value = Ticker
        EndPrice = ws.Cells(i, 6).Value
        YearlyChange = EndPrice - StartPrice
        ws.Cells(printcount, 10).Value = YearlyChange
        If StartPrice <> 0 Then
            PercentChange = YearlyChange / StartPrice
        Else
            PercentChange = YearlyChange
        End If
        ws.Cells(printcount, 11).Value = PercentChange
        ws.Cells(printcount, 11).NumberFormat = "0.00%"
        If PercentChange > CurrentMax Then
            ws.Cells(2, 16).Value = Ticker
            ws.Cells(2, 17).Value = PercentChange
            ws.Cells(2, 17).NumberFormat = "0.00%"
            CurrentMax = PercentChange
            End If
    If PercentChange < CurrentMin Then
            ws.Cells(3, 16).Value = Ticker
            ws.Cells(3, 17).Value = PercentChange
            ws.Cells(3, 17).NumberFormat = "0.00%"
            CurrentMin = PercentChange
            End If
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        ws.Cells(printcount, 12).Value = TotalStockVolume
        If TotalStockVolume > MaxVol Then
        ws.Cells(4, 16).Value = Ticker
            ws.Cells(4, 17).Value = TotalStockVolume
            MaxVol = TotalStockVolume
        End If
        printcount = printcount + 1
        TotalStockVolume = 0
    Else
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    End If
Next i
CurrentMax = 0
CurrentMin = 0
MaxVol = 0
finalrowyearlychange = ws.Cells(Rows.Count, "J").End(xlUp).Row
For i = 2 To finalrowyearlychange
    If ws.Cells(i, 10) > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
Next
End Sub
