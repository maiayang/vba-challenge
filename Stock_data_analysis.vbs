Sub stock_summary()

Dim ws_count As Integer
Dim j As Integer
ws_count = ActiveWorkbook.Worksheets.Count

For j = 1 To ws_count
Worksheets(j).Activate

    Dim ticker_symbol As String
    Dim year_open As Double
    Dim year_close As Double
    Dim total_volume As Double
    total_volume = 0
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim summary_table_row As Long
    summary_table_row = 2
    year_open = 1
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To last_row
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            year_open = Cells(i, 3).Value
            total_volume = (total_volume + Cells(i, 7).Value)
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_symbol = Cells(i, 1).Value
            Range("I" & summary_table_row).Value = ticker_symbol

            total_volume = total_volume + Cells(i, 7).Value
            Range("L" & summary_table_row).Value = total_volume
            
            year_close = Cells(i, 6).Value
            Range("K" & summary_table_row).Value = (Round((((year_close - year_open) / year_open) * 100), 2) & "%")
            Range("J" & summary_table_row).Value = year_close - year_open
                If Range("J" & summary_table_row).Value > 0 Then
                    Range("J" & summary_table_row).Interior.ColorIndex = 4
                ElseIf Range("J" & summary_table_row).Value < 0 Then
                    Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
            summary_table_row = summary_table_row + 1
            total_volume = 0
        Else
            total_volume = (total_volume + Cells(i, 7).Value)
        End If
    Next i

    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    Dim last_row_summary As Long
    last_row_summary = Cells(Rows.Count, 11).End(xlUp).Row
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total As Double
    
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total = 0
    
    For i = 2 To last_row_summary
        If Cells(i, 11).Value > greatest_increase Then
            greatest_increase = Cells(i, 11).Value
            Range("O2").Value = Cells(i, 9).Value
            Range("P2").Value = ((Cells(i, 11).Value) * 100 & "%")
        End If
    Next i
    
    For i = 2 To last_row_summary
        If Cells(i, 11).Value < greatest_decrease Then
            greatest_decrease = Cells(i, 11).Value
            Range("O3").Value = Cells(i, 9).Value
            Range("P3").Value = ((Cells(i, 11).Value) * 100 & "%")
        End If
    Next i
    
    For i = 2 To last_row_summary
        If Cells(i, 12).Value > greatest_total Then
            greatest_total = Cells(i, 12).Value
            Range("O4").Value = Cells(i, 9).Value
            Range("P4").Value = Cells(i, 12).Value
        End If
    Next i

Worksheets(j).Columns("A:P").AutoFit

Next j

End Sub
