Attribute VB_Name = "Module1"
Sub hwpart()

Dim counter As Integer
Dim lastrow As Long
Dim open_counter As Long
Dim open_value As Double
Dim close_value As Double
Dim total_vol As LongLong
Dim ws As Worksheet


For Each ws In Worksheets

    'Creates titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Value"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'Counter and last row
    counter = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    open_counter = 2

    'Creates tickers
    For i = 2 To lastrow
        
        'check next ticker and last ticker and prints last ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
            
            'selects open values
            open_value = ws.Cells(open_counter, 3).Value
            
            'selects close value
            close_value = ws.Cells(i, 6).Value
            
            'prints difference between close and open for given year i.e. yearly change [Column J]
            ws.Cells(counter, 10).Value = (close_value - open_value)
            
            'prints yearly change percentage [Column K], this one is for the scenario that opener is 0.
            If (open_value = 0 And ws.Cells(counter, 10).Value = 0) Or (open_value = 0 And ws.Cells(counter, 10).Value <> 0) Then
                open_value = 1
                ws.Cells(counter, 11).Value = Format((ws.Cells(counter, 10).Value / (open_value)), "Percent")
                ws.Cells(counter, 11).Interior.ColorIndex = 6
                
            Else
                ws.Cells(counter, 11).Value = Format((ws.Cells(counter, 10).Value / (open_value)), "Percent")
                    
                    'formats colors in cells (green or red)
                    If ws.Cells(counter, 10).Value > 0 Then
                        ws.Cells(counter, 10).Interior.ColorIndex = 4
                        
                    Else
                        ws.Cells(counter, 10).Interior.ColorIndex = 3
            
                    End If
            
            End If
            
            'sums the total stock value [Column L]
            total_vol = WorksheetFunction.Sum(ws.Range("G" & open_counter & ":G" & i))
            ws.Cells(counter, 12).Value = total_vol
            
            'counter for next row
            counter = counter + 1
            
            'counter for open value for each ticker
            open_counter = i + 1
            
        End If
        
    Next i
    
    'prints maxs and min for column P
    ws.Cells(2, 16).Value = Format(Application.WorksheetFunction.Max(ws.Range("K:K")), "Percent")
    ws.Cells(3, 16).Value = Format(Application.WorksheetFunction.Min(ws.Range("K:K")), "Percent")
    ws.Cells(4, 16).Value = Format(Application.WorksheetFunction.Max(ws.Range("L:L")), "Scientific")
    
    lastrow_k = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'if statement checks and selects min and max for row in column (N) and returns value.
    For i = 2 To lastrow_k
        
        If ws.Cells(i, 11).Value = ws.Cells(2, 16).Value Then
        ws.Range("O2").Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 16).Value Then
        ws.Range("O3").Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 16).Value Then
        ws.Range("O4").Value = ws.Cells(i, 9).Value
        
        End If
        
    Next i
    
    'Autofit columns
    ws.Cells.EntireColumn.AutoFit

Next ws

End Sub

