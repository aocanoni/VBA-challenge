Attribute VB_Name = "Module1"
Sub multiple_year_stock_data()

    'loop through all worksheets
    For Each ws In Worksheets
        

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ticker_total = 0
        
        summary_table_row = 2
        
        ticker_open = ws.Cells(2, 3).Value
        
        ticker_close = 0
    
    
        'loop through all tickers
        For i = 2 To LastRow
        
            'check to see if we're still within same ticker, if not then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'set ticker name
                ticker_name = ws.Cells(i, 1).Value
                
                'set ticker total
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                
                'set ticker close value
                ticker_close = ticker_close + ws.Cells(i, 6).Value
                
                'print Ticker title
                ws.Range("I1").Value = "Ticker"
                
                'print ticker name into table
                ws.Range("I" & summary_table_row).Value = ticker_name
                
                'print Total Stock Volume title
                ws.Range("L1").Value = "Total Stock Volume"
                
                'print ticker amount into table
                ws.Range("L" & summary_table_row).Value = ticker_total
                
                'set yearly change
                ticker_yearly_change = (ticker_close - ticker_open)
                
                'set percent change
                ticker_percent_change = (ticker_yearly_change / ticker_open)
                
                'print Yearly Change title
                ws.Range("J1").Value = "Yearly Change"
                
                'print yearly change into table
                ws.Range("J" & summary_table_row).Value = ticker_yearly_change
                
                'convert into currency
                ws.Range("J" & summary_table_row).NumberFormat = "$#,##0.00"
                
                'print Percentage Change title
                ws.Range("K1").Value = "Percentage Change"
                
                'print percent change into table
                ws.Range("K" & summary_table_row).Value = ticker_percent_change
                
                'convert into percent
                ws.Range("K" & summary_table_row).Style = "Percent"
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                
                'add one to sum table
                summary_table_row = summary_table_row + 1
                
                'reset total to 0
                ticker_total = 0
                
                'set ticker_open
                ticker_open = ws.Cells(i + 1, 3).Value
                
                ticker_close = 0
                
                ticker_yearly_change = 0
                
                ticker_percent_change = 0


            Else
            
                'add to ticker total
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                
            End If
            
        Next i
    
    Next ws
    
    yearly_change_color
    
    ticker_greatest
     
    
End Sub

Sub yearly_change_color()

    For Each ws In Worksheets
    
        TableLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        
        For i = 2 To TableLastRow

            'color yearly change boxes to red if negative and green if positive
            If ws.Cells(i, "J").Value >= 0 Then
                ws.Cells(i, "J").Interior.ColorIndex = 4
                    
            Else
                ws.Cells(i, "J").Interior.ColorIndex = 3
            
            End If
            
        Next i
        
    Next ws
    
End Sub

Sub ticker_greatest()

    For Each ws In Worksheets
    
    TableLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
    'print ticker title
    ws.Range("P1").Value = "Ticker"
            
    'print value title
    ws.Range("Q1").Value = "Value"
            
    'print greatest increase title
    ws.Range("O2").Value = "Greatest % Increase"
            
    'print greatest decrease title
    ws.Range("O3").Value = "Greatest % Decrease"
            
    'print greatest total stock volume title
    ws.Range("O4").Value = "Greatest Total Volume"
    
    great_inc = ws.Cells(2, "K").Value
    inc_ticker = 0
    
    great_dec = ws.Cells(2, "K").Value
    dec_ticker = 0
    
    great_total_vol = ws.Cells(2, "L").Value
    vol_ticker = 0
    
        For i = 2 To TableLastRow
        
            'find greatest increase and corresponding ticker
            'if great_inc is less than or equal to ... then..
            If (great_inc <= ws.Cells(i, "K").Value) Then
                
                'inc_ticker =
                inc_ticker = ws.Cells(i, "I").Value
                
                'rewrite/save value
                great_inc = ws.Cells(i, "K").Value
                
                'print ticker name
                ws.Range("P2").Value = inc_ticker
   
            End If
            
            'find greatest decrease and corresponding ticker
            If (great_dec >= ws.Cells(i, "K").Value) Then
            
                dec_ticker = ws.Cells(i, "I").Value
                
                great_dec = ws.Cells(i, "K").Value
                
                'print ticker name
                
                ws.Range("P3").Value = dec_ticker
                
            End If
            
            'find greatest total volume and corresponding ticker
            If (great_total_vol <= ws.Cells(i, "L").Value) Then
            
                vol_ticker = ws.Cells(i, "I").Value
                
                great_total_vol = ws.Cells(i, "L").Value
                
                'print ticker name
                ws.Range("P4").Value = vol_ticker
            
            End If
            
            
        Next i
        
                'print greatest increase and format
                ws.Range("Q2").Value = great_inc
                ws.Range("Q2").Style = "Percent"
                ws.Range("Q2").NumberFormat = "0.00%"
                              
                'print greatest decrease and format
                ws.Range("Q3").Value = great_dec
                ws.Range("Q3").Style = "Percent"
                ws.Range("Q3").NumberFormat = "0.00%"
                
                'print greatest total stock volume
                ws.Range("Q4").Value = great_total_vol
                
                
    ws.Columns("A:Q").AutoFit
    
    Next ws
    
End Sub

