
Sub addTicker():
   For Each ws In Worksheets
    
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Double
    Dim RowCount As Long
    

    YearlyOpen = ws.Cells(2, 3).Value
    YearlyClose = 0
    PercentChange = 0
    TotalVolume = 0
    SummaryRow = 2
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To RowCount
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Range("I" & SummaryRow).Value = ws.Cells(i, 1).Value
    
            'Calculate yearly change
       YearlyClose = ws.Cells(i, 6).Value
       YearlyChange = (YearlyClose - YearlyOpen)
        ws.Range("J" & SummaryRow).Value = YearlyChange
       'Conditional Formatting for Yearly Change
       If YearlyChange > 0 Then
       ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
       Else
       ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
       End If

       

        If (YearlyOpen = 0 And YearlyClose = 0) Then
           PercentChange = 0
       ElseIf (YearlyOpen = 0 And YearlyClose <> 0) Then
           PercentChange = 1
       Else
           PercentChange = (YearlyChange / YearlyOpen)
           ws.Range("K" & SummaryRow).Value = PercentChange
           ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
       End If


'Calculate total stock volume
       TotalVolume = TotalVolume + ws.Cells(i, 7).Value
       ws.Range("L" & SummaryRow).Value = TotalVolume
       SummaryRow = SummaryRow + 1
       'reset stock volume & open price
       TotalVolume = 0
       YearlyOpen = ws.Cells(i + 1, 3).Value
       
        Else
       TotalVolume = TotalVolume + ws.Cells(i, 7)


    End If
        
    Next i
    Next ws
End Sub
