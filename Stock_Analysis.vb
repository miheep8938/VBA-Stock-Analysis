Sub Stockdata()

'loop through all sheets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    'variables
    Dim open_price As Double
    open_price = 0
    
    Dim close_price As Double
    close_price = 0
    
    Dim yearly_change As Double
    yearly_change = 0
    
    Dim ticker_symbol As String
    
    Dim percent_change As Double
    percent_change = 0
    
    Dim sum_volume As Double
    sum_volume = 0
    
    Dim summary_index As Integer
    summary_index = 2
    
    'headings
    ws.Cells(1, 9).Value = "Tickers"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Determine lastrow
    Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).row
    
    'loop
    For i = 2 To Lastrow
        
        'conditional formatting for the open price
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
            'set the open price
            open_price = ws.Cells(i, 3).Value
            
        End If
        
            
        'conditional conditional formatting for the rest of the values
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'calculate sum of stock volume
            sum_volume = sum_volume + ws.Cells(i, 7).Value
        
            'Print total volume in the summary_index
            ws.Cells(summary_index, 12).Value = sum_volume
            
            'Print ticker symbol in the summary_index
            ws.Cells(summary_index, 9).Value = ws.Cells(i, 1).Value
        
            'Calculate the Price Change
            close_price = ws.Cells(i, 6).Value
           yearly_change = close_price - open_price
            
            'Print the price change in the summary index
            ws.Cells(summary_index, 10).Value = yearly_change
            
                'conditional formatting to highlight postive or negative change
                If yearly_change >= 0 Then
                    ws.Cells(summary_index, 10).Interior.ColorIndex = 4
                
                Else
                    ws.Cells(summary_index, 10).Interior.ColorIndex = 3
                    
                End If
                
            'Conditional formatting for calculating percent change
            If open_price = 0 And close_price = 0 Then
                'when open price and close price are zero, it would be 0
                percent_change = 0
                ws.Cells(summary_index, 11).Value = percent_change
                ws.Cells(summary_index, 11).Value = NumberFormat = "0.00%"
                
            ElseIf open_price = 0 Then
                'If a stock starts at zero then increases,the number would be infinite.
                'label as "New Stock" for these tickers
                Dim percent_change_NS As String
                percent_change_NS = "New Stock"
                ws.Cells(summary_index, 11).Value = percent_change
                
            Else
                percent_change = yearly_change / open_price
                ws.Cells(summary_index, 11).Value = percent_change
                ws.Cells(summary_index, 11).NumberFormat = "0.00%"
            End If
                
        'reset count and volume
        summary_index = summary_index + 1
        sum_volume = 0
        close_price = 0
        open_price = 0
        year_change = 0
        
        
        Else
        
            sum_volume = sum_volume + ws.Cells(i, 7).Value
            
        End If
    Next i

    'summary best and worst performance in a year
    'headings
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Assign lastrow of the summary table
    Summary_lastrow = ws.Cells(Rows.Count, "I").End(xlUp).row
    
    'Find the greatest and lowest percent changes as well as the greatest total volume in the summary table
For i = 2 To Summary_lastrow
    'greatest percent change
    If ws.Cells(i, 11).Value = ws.Application.WorksheetFunction.Max(Range("K2:K" & Summary_lastrow)) Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
            
        'lowest percent change
    ElseIf ws.Cells(i, 11).Value = ws.Application.WorksheetFunction.Min(Range("K2:K" & Summary_lastrow)) Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
            
        'greatest stockvolume
    ElseIf ws.Cells(i, 12).Value = ws.Application.WorksheetFunction.Max(Range("L2:L" & Summary_lastrow)) Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        
        End If
    Next i
    


Next ws


End Sub
