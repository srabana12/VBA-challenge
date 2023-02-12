Attribute VB_Name = "Module1"
Option Explicit
Sub stock_market()

    
    Dim ticker_curr, ticker_inc, ticker_dec, ticker_greatest_vol As String
    Dim total_stock, open_price, close_price, yearly_change, change_ratio, greatest_inc_rat, greatest_dec_rat, greatest_total_vol As Double
    Dim lastrow, i, j As Long
    Dim ws As Worksheet

    
    'Code to loop through each worksheet
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        'get the last row for each worksheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
            'Define the header for the new columns created during calculation
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Gretest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            j = 2
            open_price = ws.Cells(2, 3).Value
            
            'loop through all the existing rows to find unique ticker and calculate associated values
            For i = 2 To lastrow
                ticker_curr = ws.Cells(i, 1).Value
                
                'check if the next row is not same ticker
                If ticker_curr <> ws.Cells(i + 1, 1).Value Then

                    'calculatations
                    close_price = ws.Cells(i, 6).Value
                    yearly_change = close_price - open_price
                    change_ratio = yearly_change / open_price
                    
                    If change_ratio > greatest_inc_rat Then
                        greatest_inc_rat = change_ratio
                        ticker_inc = ticker_curr
                    ElseIf change_ratio < greatest_dec_rat Then
                        greatest_dec_rat = change_ratio
                        ticker_dec = ticker_curr
                    ElseIf total_stock > greatest_total_vol Then
                        greatest_total_vol = total_stock
                        ticker_greatest_vol = ticker_curr
                    End If
                
                    'output
                    ws.Cells(j, 9).Value = ticker_curr
                    ws.Cells(j, 10).Value = yearly_change
                    ws.Cells(j, 11).Value = FormatPercent(change_ratio, 2)
                    ws.Cells(j, 12).Value = total_stock
                    
                    
                    'Conditional Formatting for yearly change
                    If yearly_change < 0 Then
                        ws.Cells(j, 10).Interior.ColorIndex = 3 'Red
                     Else
                        ws.Cells(j, 10).Interior.ColorIndex = 4  'Green
                    End If
                    
                    'Conditional Formatting for percent change ratio
                    If change_ratio < 0 Then
                        ws.Cells(j, 11).Interior.ColorIndex = 3 'Red
                     Else
                        ws.Cells(j, 11).Interior.ColorIndex = 4  'Green
                    End If

                    'prepare for next stock
                    open_price = ws.Cells(i + 1, 3).Value
                    total_stock = 0
                     j = j + 1
                     
                End If
                'keep adding total stock if same ticker
                total_stock = total_stock + ws.Cells(i, 7).Value
                
         Next i
         
         'output value for greatest increase,decrese and total volume for ticker
           ws.Cells(2, 16).Value = ticker_inc
           ws.Cells(3, 16).Value = ticker_dec
           ws.Cells(4, 16).Value = ticker_greatest_vol
           ws.Cells(2, 17).Value = FormatPercent(greatest_inc_rat, 2)
           ws.Cells(3, 17).Value = FormatPercent(greatest_dec_rat, 2)
           ws.Cells(4, 17).Value = greatest_total_vol
                   
        Next ws
End Sub

