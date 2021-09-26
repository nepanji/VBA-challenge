Attribute VB_Name = "Module1"
Sub Analysis_Stock():
    For Each ws In Worksheets
        
        'Worksheets(ws).Activate
        
        'Create initial variables needed to evaluate ticker data
        Dim i As Double
        Dim Summary_i As Double
        Dim LastRow As Double
        
        'Create variables needed to develop summary table data
        Dim open_price As Double
        Dim close_price As Double
        Dim price_change As Double
        Dim percent_price_change As Double
        Dim total_volume As Double
        
        ' Fill in the summary column names
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        
        'Fill in the summary row and column names for the second summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Set the values for open price and total volume
        open_price = 0
        total_volume = 0
        Summary_i = 1
        
        'Set variable for last row
        LastRow = ws.UsedRange.Rows.Count
        
        
        'Start loop through data table
        For i = 2 To LastRow
             
            'Set variable for ticker and first open price
            ticker = ws.Cells(i, 1).Value
            next_ticker = ws.Cells(i + 1, 1).Value
            first_price_change = ws.Cells(i, 6).Value - ws.Cells(2, 3).Value
            
            'Enter first open price in Summary Table in cell K2
            ws.Cells(2, 11).Value = first_price_change
            
            'Color the Yearly cell for the first value based on the outcome in Column K
            If first_price_change < 0 Then
                ws.Cells(2, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(2, 11).Interior.ColorIndex = 4
            End If
            
            'Compare Ticker names in Column A
            If next_ticker <> ticker Then
            
                'Enter Ticker names to the Summary Table in Column J
                ws.Cells(Summary_i + 1, 10).Value = ticker
                
                'Calculate open and closing price
                open_price = ws.Cells(i, 3).Value
                close_price = ws.Cells(i + 1, 6).Value
                
                'Calculate Yearly Change by subtracting closing price from opening price
                price_change = close_price - open_price
                
                'Enter Yearly Change to the Summary Table in Column K
                ws.Cells(Summary_i + 1, 11).Value = price_change
                
                'Color the Yearly cells based on their outcomes in Column K
                If price_change < 0 Then
                    ws.Cells(Summary_i + 1, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(Summary_i + 1, 11).Interior.ColorIndex = 4
                End If
            
                'Calculate the percent of annual change(*check for when open_price is zero)
                If open_price = 0 Then
                    percent_price_change = 0
                Else
                 percent_price_change = ((close_price - open_price) / open_price)
                End If
                
                'Enter % of change in Column L
                ws.Cells(Summary_i + 1, 12).Value = percent_price_change
                
                'Format percent_price_change cell to % in Column L
                ws.Cells(Summary_i + 1, 12).NumberFormat = ".00%"
                
                'Enter N/A when open price is zero
                ws.Cells(Summary_i + 1, 12).Value = percent_price_change
                
                
                'Calculate the Total Volume of the Ticker in Column G
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                'Enter total stock change in Column M
                ws.Cells(Summary_i + 1, 13).Value = total_volume
                
                
               
                
                Summary_i = Summary_i + 1
                total_volume = 0
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value
                
            End If
            
            
        Next i
        
    
        
        'Find the MAX value from Column L to determine greatest increase
        Set checkPercentRange = ws.Range("L:L")
        MAX_Increase = Application.WorksheetFunction.Max(checkPercentRange)
        
        'Enter greatest % increase ticker name in Row 3, Column P and the value in Row 3, Column Q
        ws.Cells(2, 16).Value = MAX_Increase
        ws.Cells(2, 17).Value = ws.Cells(Application.Match(MAX_Increase, ws.Range("L:L"), False), 10).Value
        
        
        'Format greatest % increase cell to % in Column P
        ws.Cells(2, 16).NumberFormat = ".00%"
        
        'Find the MIN value from Column L to determine greatest decrease
        Set checkPercentRange = ws.Range("L:L")
        MIN_Increase = Application.WorksheetFunction.Min(checkPercentRange)
        
        'Enter greatest % decrease ticker name in Row 2, Column P and the value in Row 2, Column Q
        ws.Cells(3, 16).Value = MIN_Increase
        ws.Cells(3, 17).Value = ws.Cells(Application.Match(MIN_Increase, ws.Range("L:L"), False), 10).Value
        
        'Format greatest % increase cell to % in Column P
        ws.Cells(3, 16).NumberFormat = ".00%"
        
        'Find the MAX value from Column M to determine greatest total volume increase
        Set checkVolumeRange = ws.Range("M:M")
        MAX_Volume_Increase = Application.WorksheetFunction.Max(checkVolumeRange)
        
        'Enter greatest % increase ticker name in Row 3, Column P and the value in Row 3, Column Q
        ws.Cells(4, 16).Value = MAX_Volume_Increase
        ws.Cells(4, 17).Value = ws.Cells(Application.Match(MAX_Volume_Increase, ws.Range("M:M"), False), 10).Value
        
        'Format greatest % increase cell to general in Column P
        ws.Cells(4, 16).NumberFormat = "0.00E+10"
        
        'AutoFit Header Column Width
        ws.Range("J1:P17").Columns.AutoFit
    
    Next ws

End Sub

