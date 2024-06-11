Sub stocks()

' Define variable to loop through worksheet
Dim ws As Worksheet

'Define date variables
Dim min_date As Date
Dim max_date As Date

Dim volume As Variant

' Define summary_row variable
Dim summary_row As Integer

    'Begin loop through worksheets
    For Each ws In Worksheets
    
        ' Create headers for unique value columns
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "quarterly_change"
        ws.Cells(1, 11).Value = "percent_change"
        ws.Cells(1, 12).Value = "total_stock_volume"
        
        ws.Cells(2, 15).Value = "greatest_percent_increase"
        ws.Cells(3, 15).Value = "greatest_percent_decrease"
        ws.Cells(4, 15).Value = "greatest_stock_volume"
        ws.Cells(1, 16).Value = "ticker"
        ws.Cells(1, 17).Value = "value"
        
        ' Find minimum and maximum dates
        min_date = Application.WorksheetFunction.Min(ws.Columns(2))
        max_date = Application.WorksheetFunction.Max(ws.Columns(2))
        
        'Find Last Row
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set variable for row of summary table
        summary_row = 2
        
            ' Begin loop through individual sheet
            For j = 1 To lastRow
            
                ' Set volume at 0
                 
                ' Searches for unique cell values
                If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                
                    ' Find ticker and opening value
                    ticker = ws.Cells(j + 1, 1).Value
           
                End If
                
                ' Searches for row for each ticker
                If ws.Cells(j + 1, 1).Value = ticker Then
                    
                    ' Adds up stock volume for that ticker
                    volume = volume + ws.Cells(j + 1, 7).Value
                    
                End If
                
                ' Find opening value based on minimum date for each ticker symbol
                If ws.Cells(j, 1).Value = ticker And ws.Cells(j, 2) = min_date Then
                
                    ' Record opening value
                    opening = ws.Cells(j, 3).Value
                
                End If
                
                ' Find closing value based on maximum date for each ticker value
                If ws.Cells(j + 1, 1).Value = ticker And ws.Cells(j + 1, 2) = max_date Then
                    
                    ' Record closing value
                    closing = ws.Cells(j + 1, 6).Value
                
                    ' Calculate quarterly change
                    quarterly_change = closing - opening
                    
                    'Calculate percentage change
                    percent_change = quarterly_change / opening
                    
                    
                    ' Insert into summary columns
                    ws.Cells(summary_row, 9).Value = ticker
                    ws.Cells(summary_row, 10).Value = quarterly_change
                    ws.Cells(summary_row, 11).Value = percent_change
                    ws.Cells(summary_row, 12).Value = volume
                     
                     ' Format percentage column
                    ws.Cells(summary_row, 11).NumberFormat = "0.00%"
                    
                    ' Reset volume for 0 for next ticker
                    volume = 0
                    
                    ' Color code quarterly change based on positive and negative values
                    If quarterly_change >= 0 Then
                        ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                    
                    End If

                    ' Move to next row of summary table
                    summary_row = summary_row + 1
                    
                End If
    
            Next j
            
            ' Find greatest percent increase, greatest percent decrease, and greatest total volume
            greatest_increase = Application.WorksheetFunction.Max(ws.Columns(11))
            greatest_decrease = Application.WorksheetFunction.Min(ws.Columns(11))
            greatest_volume = Application.WorksheetFunction.Max(ws.Columns(12))
            
            ' Put values into table
            ws.Cells(2, 17).Value = greatest_increase
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).Value = greatest_decrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(4, 17).Value = greatest_volume
            
            ' Add ticker values to table
            For i = 1 To lastRow
            
            If ws.Cells(i, 11).Value = greatest_increase Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value = greatest_decrease Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value = greatest_volume Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
            
            Next i
             

    Next ws

End Sub


