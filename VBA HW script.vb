Sub StockData()


    For Each ws In Worksheets
    
        Dim StockSheets As String
        'Current row
        Dim i As Long
        'Start row of ticker block
        Dim j As Long
        'Index counter to fill Ticker row
        Dim TickCount As Long
        'Last row column A
        Dim LastRowA As Long
        'Variable for percent change calculation
        Dim PerChange As Double
        

        'Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Start of Ticker Counter
        TickCount = 2
        j = 2
        
        'Finding the last row of each sheet
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        'Loop through all rows
        For i = 2 To LastRowA
            
            'Check to see if ticker name has changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Print ticker in Column I
            ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
            
            'Calculate Yearly Change in Column J
            ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
            
                'Conditional formating
                    If ws.Cells(TickCount, 10).Value < 0 Then
                
                    'Set cell color to red
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell color to green
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
            
                    End If
                    
                'Calculate percent change in column K
                If ws.Cells(j, 3).Value <> 0 Then
                    
                PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                'Format cells to percentages
                ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
                Else
                    
                ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                End If
            
            'Calculate and write total volume in column L (#12)
            ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
            'Increase TickCount by 1
            TickCount = TickCount + 1
                
            'Set new start row of the ticker block
            j = i + 1
            
            
            End If
            
            
        Next i


    Next ws

End Sub