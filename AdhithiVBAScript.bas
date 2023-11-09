Attribute VB_Name = "Module1"
Sub stockscreener():
    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim ticker As String
        Dim LastRow As Long
        Dim yearlychange As Double
        Dim trading_volume As Variant
        Dim close_value As Double
        Dim open_value As Double
        Dim percentchange As Double
        
                
        'Creates 4 new columns for the summary table
            'ws.Range("K1").EntireColumn.Insert
            ws.Cells(1, 11).Value = "Ticker"
        
            'ws.Range("L1").EntireColumn.Insert
            ws.Cells(1, 12).Value = "Yearly Change"
    
            'ws.Range("M1").EntireColumn.Insert
            ws.Cells(1, 13).Value = "Percentage Change"
        
            'ws.Range("N1").EntireColumn.Insert
            ws.Cells(1, 14).Value = "Total Stock Volume"
            
            ws.Cells(1, 17).Value = "Ticker"
            ws.Cells(1, 18).Value = "Value"
            ws.Cells(2, 16).Value = "Greatest % Increase"
            ws.Cells(3, 16).Value = "Greatest % Decrease"
            ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        trading_volume = 0
        
        summary_table_row = 2
         
        yearlychange = 0
        open_value = 0
        close_value = 0
    
        
        For i = 2 To LastRow
            If ws.Cells(i, 2).Value = ws.Cells(2, 2).Value Then
                open_value = ws.Cells(i, 3).Value
                trading_volume = CDec(trading_volume + ws.Cells(i, 7).Value)
        
        
            ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                trading_volume = CDec(trading_volume + ws.Cells(i, 7).Value)
                
                
            Else
                
                ticker = ws.Cells(i, 1).Value
                
                'Set the total trading volume
                trading_volume = CDec(trading_volume + ws.Cells(i, 7).Value)
                
                'Set the close value to the cell value in column 6 at the end of the for loop
                close_value = ws.Cells(i, 6).Value
                
                'Calculate the yearly change
                yearlychange = close_value - open_value
                
                'Calculate the percentage change
                percentchange = (yearlychange / (open_value))
                
                'the ticker goes into column K of the summary table
                ws.Range("K" & summary_table_row).Value = ticker
                
                'yearly change goes into column L of the summary table
                ws.Range("L" & summary_table_row).Value = yearlychange
                
                'Percentage change goes into column M of the summary table
                ws.Range("M" & summary_table_row).Value = percentchange
                ws.Range("M" & summary_table_row).NumberFormat = "0.00%"
                
                'the total stock volume goes into column N of the summary table
                ws.Range("N" & summary_table_row).Value = trading_volume
                
                'Add one more row to the summary table for the next stock
                summary_table_row = summary_table_row + 1
                
                'Reset the total trading volume for each stock
                trading_volume = 0
            
            
            End If
            
            Next i
            
            For i = 2 To 3001
            
                If ws.Cells(i, 12).Value > 0 Then
                    ws.Cells(i, 12).Interior.ColorIndex = 10
                Else
                    ws.Cells(i, 12).Interior.ColorIndex = 3
                End If
            
            Next i
            
        percentrange = ws.Range("M2:M3001")
        tickerrange = ws.Range("K2:K3001")
            
        'greatest % increase
         greatest_percent_increase = Application.WorksheetFunction.Max(percentrange)
            
            
         ws.Cells(2, 18).Value = greatest_percent_increase
         ws.Cells(2, 18).NumberFormat = "0.00%"
         
         
         'greatest % decrease
            
        greatest_percent_decrease = Application.WorksheetFunction.Min(percentrange)
        ws.Cells(3, 18).Value = greatest_percent_decrease
        ws.Cells(3, 18).NumberFormat = "0.00%"
        
        'greatest total volume
         VolRange = ws.Range("N2:N3001")
            
         greatest_total_volume = Application.WorksheetFunction.Max(VolRange)
         ws.Cells(4, 18).Value = greatest_total_volume
         ws.Cells(4, 18).NumberFormat = "General"
         
         
        'maxticker = 0
        'minticker = 0
            
        For i = 2 To 3001
        
            If ws.Cells(i, 13).Value = greatest_percent_increase Then
            
                maxticker = ws.Cells(i, 11).Value
                'The max ticker symbol goes into Column Q
                ws.Cells(2, 17).Value = maxticker
                
            ElseIf ws.Cells(i, 13).Value = greatest_percent_decrease Then
            
                minticker = ws.Cells(i, 11).Value
                'The min ticker symbol goes into column Q
                ws.Cells(3, 17).Value = minticker
            
            ElseIf ws.Cells(i, 14).Value = greatest_total_volume Then
                
                total_volume_ticker = ws.Cells(i, 11).Value
                ws.Cells(4, 17).Value = total_volume_ticker
            
                
            End If
            
            Next i
            
    
        Next ws
        




End Sub
