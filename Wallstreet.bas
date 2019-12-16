Attribute VB_Name = "Module1"
Sub Wallstreet()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Activate
        
        Dim ticker As String
        Dim year_opener As Double
        Dim year_closer As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        ' Dim total_volume As Long
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
            
            ' Dim volume As Long
            volume = Cells(i, 7).Value
            total_volume = total_volume + volume
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ticker = Cells(i, 1).Value
                year_closer = Cells(i, 6).Value
                yearly_change = year_closer - year_opener
                
                If yearly_change = 0 Then
                    
                    percent_change = 0
                
                ElseIf year_opener = 0 Then
                    
                    percent_change = yearly_change
                
                Else
                
                    percent_change = yearly_change / year_opener
                
                End If
                
                Cells(summary_row, 9).Value = ticker
                Cells(summary_row, 10).Value = yearly_change
                Cells(summary_row, 11).Value = percent_change
                Cells(summary_row, 11).NumberFormat = "0.00%"
                Cells(summary_row, 12).Value = total_volume
                
                If Cells(summary_row, 11) > 0 Then
                    
                    Cells(summary_row, 11).Interior.ColorIndex = 4
                    
                ElseIf Cells(summary_row, 11) < 0 Then
                    
                    Cells(summary_row, 11).Interior.ColorIndex = 3
                
                End If
                
                year_opener = Cells(i + 1, 3).Value
                total_volume = 0
                summary_row = summary_row + 1
                
            End If
        Next i
        
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
        gretest_hits = 2
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0
        
        For i = 2 To last_row
            
            If Cells(i, 11) > greatest_increase Then
                
                greatest_increase = Cells(i, 11)
                Range("P2") = Cells(i, 9)
                Range("Q2") = greatest_increase
                Range("Q2").NumberFormat = "0.00%"
                
            End If
            
            If Cells(i, 11) < greatest_decrease Then
                
                greatest_decrease = Cells(i, 11)
                Range("P3") = Cells(i, 9)
                Range("Q3") = greatest_decrease
                Range("Q3").NumberFormat = "0.00%"
                
            End If
            
            If Cells(i, 12) > greatest_volume Then
                
                greatest_volume = Cells(i, 12)
                Range("P4") = Cells(i, 9)
                Range("Q4") = greatest_volume
                
                
            End If
            
        Next i
        
        Columns("I:Q").AutoFit
    
    Next ws

End Sub


