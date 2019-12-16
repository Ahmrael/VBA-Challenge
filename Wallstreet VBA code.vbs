Sub Wallstreet()
    
    For Each ws In Worksheets
        
        Dim ticker As String
        Dim year_opener As Double
        Dim year_closer As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_volume As Long
        Dim summary_row As Integer
        Dim last_row As Long
        
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        
        last_row = Range("A" & Rows.Count).End(xlUp).Row
        
        summary_row = 2
        total_volume = 0
        year_opener = Cells(2, 3).Value
        
        For i = 2 To last_row
            
            Dim volume as Long
            volume = Cells(i, 7).Value
            total_volume = total_volume + volume
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ticker = Cells(i, 1).Value
                year_closer = Cells(i, 6).Value
                yearly_change = year_closer - year_opener
                percent_change = yearly_change / year_opener
                
                Cells(summary_row, 9).Value = ticker
                Cells(summary_row, 10).Value = yearly_change
                Cells(summary_row, 11).Value = percent_change
                Cells(summary_row, 11).NumberFormat = "0.00%"
                Cells(summary_row, 12).Value = totale_volume
                
                year_opener = Cells(i + 1, 3).Value
                total_volume = 0
                summary_row = summary_row + 1
                
            End If
        Next i
    Next ws

End Sub
