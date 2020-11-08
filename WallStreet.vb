Sub wallStreet():

'Assign Variables

    Dim Volume As Double
    Dim Total As Double
    Dim Yearly As Double
    Dim Percent As Double
    Dim J As Long
    Dim ws_Count As Integer
    Dim L As Integer
   
For Each ws In Worksheets
    J = 2

'Create column headers

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % of Increase"
    ws.Cells(3, 14).Value = "Greatest % of Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"

'Pull Tickers into the Ticker column
    
    RCount = 2
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        For I = 2 To RowCount
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                Total = Total + ws.Cells(I, 7).Value
                If ws.Cells(J, 3) = 0 Then
                    For nonZeroValue = J To I
                        If ws.Cells(nonZeroValue, 3).Value <> 0 Then
                            J = nonZeroValue
                            Exit For
                        End If
                    Next nonZeroValue
                End If
                 
'Pull totals to appropriate columns
                
                Yearly = ws.Cells(I, 6).Value - ws.Cells(J, 3).Value
                Percent = (Yearly / ws.Cells(J, 3).Value) * 100
                'J = J + 1
                ws.Range("I" & RCount).Value = ws.Cells(I, 1).Value
                ws.Range("J" & RCount).Value = Yearly
                ws.Range("J" & RCount).Style = Yearly
                ws.Range("K" & RCount).Value = Round(Percent, 2)
                ws.Range("L" & RCount).Value = Total
            
'Shade Positve cells green and negative cells red
                
                If Yearly > 0 Then
                    ws.Range("J" & RCount).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & RCount).Interior.ColorIndex = 3
                End If
                
'Reset counters for next loop
                
                Total = 0
                RCount = RCount + 1
                Yearly = 0
            Else
                Total = Total + ws.Cells(I, 7).Value
            End If
        Next I
        
    SumRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Find greatest increase, decrease and volume

    Set rng = ws.Range("K2:K" & SumRow)
    Set rng2 = ws.Range("L2:L" & SumRow)
        ws.Range("P2").Value = FormatPercent(Application.WorksheetFunction.Max(rng), 2)
        ws.Range("P3").Value = FormatPercent(Application.WorksheetFunction.Min(rng), 2)
        ws.Range("P4").Value = Application.WorksheetFunction.Max(rng2)
        ws.Range("O2").Value = ws.Cells(Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(rng), rng, 0) + 1, 9).Value
        ws.Range("O3").Value = ws.Cells(Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(rng), rng, 0) + 1, 9).Value
        ws.Range("O4").Value = ws.Cells(Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(rng2), rng2, 0) + 1, 9).Value
    Next ws
    
End Sub