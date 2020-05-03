Sub Stock_Analysis()
'set ws as a variable to hold file name
Dim ws As Worksheet
'loop through all of the ws in the active workbook
For Each ws In Worksheets
    
    'set an initial variable for holding the stock symbol name
    Dim Ticker_name As String
    'set an initial variable for holding the total stock volume per ticker_name
    Dim Total_ticker_volume As Double
    Total_ticker_volume = 0
    'set variable to keep track of the location for each stock brand in the summary table
    Dim Summary_table_row As Long
    Summary_table_row = 2
    'place headers to the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    'initialize row count
    Dim Lastrow As Long
    Dim i As Long
    'determine the last row
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'initiate variable to hold open price
    Dim Open_price As Double
    'iterator
    Open_price = ws.Cells(2, 3).Value
    Dim Close_price As Double
    
    'loop through all tickers on current worksheet
    For i = 2 To Lastrow
        
        'check if we are still withing the same ticker then do what we plan
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'place ticker  name in the summary table, ticker name change
            
            Ticker_name = ws.Cells(i, 1).Value
            'calculate the difference in open-close prices
            Close_price = ws.Cells(i, 6).Value
            
            Dim Yearly_price_change As Double
            
            Yearly_price_change = Close_price - Open_price
            'initiate and print Percent_change in the summary table
            
            Dim Percent_change As Double
            'division by zero............................
            If Open_price <> 0 Then
                Percent_change = (Yearly_price_change / Open_price) * 100
            Else
                Percent_change = 0
            End If
            'open price value to start..adviced to add inside the loop
            Open_price = ws.Cells(i + 1, 3).Value
        
            'add to the Total_stock_volume
            Total_ticker_volume = Total_ticker_volume + ws.Cells(i, 7).Value
            'print the ticker name in the summary table
            ws.Range("I" & Summary_table_row).Value = Ticker_name
            'print Yearly_change of stock price in the summary table..
            ws.Range("J" & Summary_table_row).Value = Round(Yearly_price_change, 2)
            'conditional formatting that will highlight positive change in green and negative change in red
        
            If (Yearly_price_change) > 0 Then
                
                ws.Range("J" & Summary_table_row).Interior.ColorIndex = 4

            ElseIf (Yearly_price_change <= 0) Then

                ws.Range("J" & Summary_table_row).Interior.ColorIndex = 3
            End If
            
            Total_ticker_volume = Total_ticker_volume + ws.Cells(i, 7).Value
            
            ws.Range("K" & Summary_table_row).Value = (CStr(Percent_change) & "%")
            ws.Range("L" & Summary_table_row).Value = Total_ticker_volume
            'add one to the Summary_table_row to move down
            Summary_table_row = Summary_table_row + 1
            'reset the Total_stock_volume so next group of tickers added to previous
            'Total_ticker_volume = 0
            Yearly_price_change = 0
            Close_price = 0
            Open_price = ws.Cells(i + 1, 3).Value
            '.............
            Total_ticker_volume = 0
            Percent_change = 0

        Else
            
            Total_ticker_volume = Total_ticker_volume + ws.Cells(i, 7).Value

        End If
        'initiate variables to hold values for O, P, Q columns of second part of summary table
        'challenges part
        Dim Max_percent_inc As Double
        Dim Min_percent_dec As Double
        Dim Max_total_volume As Double
        Dim Ticker_colP As String
        Dim Value As Long
        'greatest percent increase and decrease, greatest total volume
        'If (Percent_change > Max_percent_inc) Then
        'Max_percent_inc = Yearly_price_change
        'Ticker_colP = Ticker_name
        'ElseIf (Percent_change < Min_percent_dec) Then
        'Min_percent_dec = Percent_change
        'Ticker_colP = Ticker_name
        'End If

    Next i

Next ws
End Sub
