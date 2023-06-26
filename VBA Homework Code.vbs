Sub MultYearStock()

    For Each ws In Worksheets
    
        'Set up variables
        Dim summary_table_row As Long
        Dim lastrow As Long
        Dim stock_volume As Double
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim greatest_percent_increase As Double
        Dim greatest_ticker As String
        Dim greastest_percent_decrease As Double
        Dim lowest_ticker As String
        Dim greatest_volume_total As Double
        Dim greatest_volume_ticker As String
        
        'To add last row: lastrow = Cells(Rows.Count,1).End(xlUp).Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'set up varibables for the total volume and greatest decrease/increase
        greatest_volume_total = 0
        greatest_volume_ticker = ""
        
        greatest_percent_increase = 0
        greatest_ticker = ""

        greatest_percent_decrease = 999999999999
        lowest_ticker = ""

        'Add summary table starting at row 2
        summary_table_row = 2
        
        'Obtain the total stock volumes
        stock_volume = 0
        
        opening_price = ws.Cells(2, 3).Value
        
        'Set up loop
        For I = 2 To lastrow
        
            stock_volume = stock_volume + ws.Cells(I, 7).Value
            
            'Use if statement to find each ticker to add to column "I"="9"
            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                
                ws.Cells(summary_table_row, 9).Value = ws.Cells(I, 1).Value
                
                'Add the sum of stock volume to summary table
                ws.Cells(summary_table_row, 12).Value = stock_volume
                
                'Calculate yearly change
                closing_price = ws.Cells(I, 6).Value
                yearly_change = closing_price - opening_price
                
                'Calculate the percentage change
                percent_change = (closing_price - opening_price) / opening_price
                
                'Add percent change to summary table
                ws.Cells(summary_table_row, 11).Value = percent_change
                
                'Add yearly change to summary table
                ws.Cells(summary_table_row, 10).Value = yearly_change
                
                'Opening mext tickers next cells
                opening_price = ws.Cells(I + 1, 3).Value
                
                'Syntax to make sure rows are being used subsequently
                summary_table_row = summary_table_row + 1

                'Find greatest total volume
                If stock_volume > greatest_volume_total Then
                    greatest_volume_total = stock_volume
                    greatest_volume_ticker = ws.Cells(I, 1).Value
                End If
            
                'Find the greatest percent increase
                If percent_change > greatest_percent_increase Then
                    greatest_percent_increase = percent_change
                    greatest_ticker = ws.Cells(I, 1).Value
                End If

                ' Find the greatest percent decrease
                If percent_change < greatest_percent_decrease Then
                    greatest_percent_decrease = percent_change
                    lowest_ticker = ws.Cells(I, 1).Value
                End If
                
                'reset every time it adds new ticker values
                stock_volume = 0
                
                'Set cell color based on percent change
                If percent_change < 0 Then
                    ws.Cells(summary_table_row - 1, 10).Interior.ColorIndex = 3 'Red color
                    
                ElseIf percent_change > 0 Then
                    ws.Cells(summary_table_row - 1, 10).Interior.ColorIndex = 4 'Green color
                End If
                
            End If
            
        Next I

        'put the greatest volume ticker in the summary table
        ws.Cells(4, 17).Value = greatest_volume_total
        ws.Cells(4, 16).Value = greatest_volume_ticker
        'Add greatest percent increase to the summary table
        ws.Cells(2, 17).Value = greatest_percent_increase
        ws.Cells(2, 16).Value = greatest_ticker
        'Add greatest percent decrease to summary table
        ws.Cells(3, 17).Value = greatest_percent_decrease
        ws.Cells(3, 16).Value = lowest_ticker
  
    Next ws

End Sub



