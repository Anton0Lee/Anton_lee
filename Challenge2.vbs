Attribute VB_Name = "Module1"
Sub stocks()
Dim ws As Worksheet
For Each ws In Worksheets
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summary_row = 2
        total_volume = 0
        ws.Range("K:K").NumberFormat = "0.00%"
            For i = 2 To lastrow
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                year_open = ws.Cells(i, 3).Value
                End If
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    Ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & summary_row).Value = Ticker
                    total_volume = total_volume + ws.Cells(i, 7).Value
                    ws.Range("L" & summary_row).Value = total_volume
                    total_volume = 0
                    year_close = ws.Cells(i, 6).Value
                    ws.Range("J" & summary_row).Value = year_close - year_open
                    ws.Range("K" & summary_row).Value = (year_close - year_open) / year_open
                    summary_row = summary_row + 1
                        Else
                        total_volume = total_volume + ws.Cells(i, 7).Value
                    End If
            Next i
        ws.Range("J:J").FormatConditions.Delete
        ws.Range("J2:J" & summary_row - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:=0).Interior.ColorIndex = 4
        ws.Range("J2:J" & summary_row - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=0).Interior.ColorIndex = 3
        ws.Range("K2:K" & summary_row - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:=0).Interior.ColorIndex = 4
        ws.Range("K2:K" & summary_row - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=0).Interior.ColorIndex = 3
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0.00E+00"
        percent_change = ws.Range("K:K")
        tickers = ws.Range("I:I")
        total_volume = ws.Range("L:L")
            ws.Range("Q2") = Application.WorksheetFunction.max(percent_change)
            ws.Range("Q3") = Application.WorksheetFunction.Min(percent_change)
            ws.Range("Q4") = Application.WorksheetFunction.max(total_volume)
                max_p = ws.Range("Q2").Value
                min_p = ws.Range("Q3").Value
                max_v = ws.Range("Q4").Value
                    ws.Range("P2") = Application.WorksheetFunction.XLookup(max_p, percent_change, tickers)
                    ws.Range("P3") = Application.WorksheetFunction.XLookup(min_p, percent_change, tickers)
                    ws.Range("P4") = Application.WorksheetFunction.XLookup(max_v, total_volume, tickers)

    ws.Columns("A:Q").AutoFit
Next ws

End Sub

