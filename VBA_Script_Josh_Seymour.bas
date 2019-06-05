Attribute VB_Name = "Module1"
Sub Stock_Market():
    'Declare ticker summary table variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim stock_volume As Double
    Dim ticker_column_row As String
        
    'Yealy change variables
    Dim year_change As Double
    Dim year_change_start As Double
    Dim year_change_end As Double
    
    'Percent change variable
    Dim percent_change As Double
       
    'Values to find the final rows for columns A, J and K
    Dim finalrow As Double
    Dim finalchange As Double
    
    'run through each worksheet and find final rows
    For Each ws In ThisWorkbook.Worksheets
    finalrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    finalchange = ws.Range("J" & ws.Rows.Count).End(xlUp).Row
    
    'Declare bonus values and set to 0 so they are cleared out before the next worksheet
    Dim greatest_increase As Double
    greatest_increase = 0
    Dim greatest_decrease As Double
    greatest_decrease = 0
    Dim greatest_volume As Double
    greatest_volume = 0
    
    'Set column headers for each worksheet
    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Year Change"
    ws.Cells(1, 10).Value = "Percent Change"
    ws.Cells(1, 11).Value = "Total Stock Volume"
    ws.Cells(1, 13).Value = "Greatest % Increase"
    ws.Cells(2, 13).Value = "Greatest % Decrease"
    ws.Cells(3, 13).Value = "Greatest Total Volume"
       
    'set year column format to percentage
    ws.Columns("J").NumberFormat = "0.00%"
    ws.Cells(1, 14).NumberFormat = "0.00%"
    ws.Cells(2, 14).NumberFormat = "0.00%"
    ws.Cells(3, 14).NumberFormat = "General"
    
    'Set-up variables to be inserted into columns
    ticker_column_row = 2
            
        'Loop through all tickers
        For i = 2 To finalrow
        
            'Capture the first value in open column to compare to last in close column
            If i = 2 Then
                year_change_start = ws.Cells(i, 3).Value
            End If
                            
            'Check if still counting same ticker, it not sum the values and add to columns
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Grab the last close value before the new ticker
                year_change_end = ws.Cells(i, 6).Value
        
                'Add ticker to Ticker column
                ticker = ws.Cells(i, 1).Value
                            
                'Calc yearly change
                year_change = year_change_end - year_change_start
                
                'Calc percent change and remove Overflow issue by not including values where the demoninator is 0
                If year_change_start = 0# Then
                    percent_change = 0
                Else
                    percent_change = (year_change / year_change_start)
                End If
                
                'If percent change is greater than the value in the greatest increase, then replace with new value
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                End If
                
                'If percent change is less than the value in the greatest decrease, then replace with new value
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                End If
                
                'Add to stock volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
                'If stock volume is greater than the greatest volume, capture the new value
                If stock_volume > greatest_volume Then
                    greatest_volume = stock_volume
                End If
                 
                'Print ticker name in Ticker Column
                ws.Range("H" & ticker_column_row).Value = ticker
                 
                'Print year change
                ws.Range("I" & ticker_column_row).Value = year_change
                 
                'print percent change
                ws.Range("J" & ticker_column_row).Value = percent_change
                
                'Print the Stock Volume to the Volume Column
                ws.Range("K" & ticker_column_row).Value = stock_volume
                
                'Add one to the ticker summary
                ticker_column_row = ticker_column_row + 1
        
                'Save year change start for next ticker
                year_change_start = ws.Cells(i + 1, 3).Value
                
                'Reset the stock volume to 0
                stock_volume = 0
                
                
            'If the cell following the prior is the same, keep counting
             Else
    
                'Add to the stock volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
        
            End If

        'Go to next row in ticker summary
        Next i
        
        'Loop to apply conditional formatting
        For K = 2 To finalchange
           If ws.Cells(K, 10) >= 0 Then
              ws.Cells(K, 10).Interior.ColorIndex = 4
           Else
              ws.Cells(K, 10).Interior.ColorIndex = 3
           End If
        Next K
        
        'Write greatest increase, decrease and volume to worksheet before going to the next one
        ws.Cells(1, 14).Value = greatest_increase
        ws.Cells(2, 14).Value = greatest_decrease
        ws.Cells(3, 14).Value = greatest_volume
        
    'move to next worksheet
     Next ws
    
End Sub



