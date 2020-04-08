Attribute VB_Name = "Module1"
Sub StockData()
    
    ' Loop through all active sheets
    
    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

    'Create heading for first summary
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
        
    'Create Variable to hold value
    Dim Price_Open As Double
    Dim Price_Close As Double
    Dim Yearly_Change As Double
    Dim Ticker_Symbol As String
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim Summary_Table As Double
        
    Summary_Table = 2
        
    Total_Volume = 0
        
    'Determine the the Last Row
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Identify the stock open price
    Price_Open = Cells(2, 3).Value
        
    'Loop through all ticker symbol, and yearly stock change
    For i = 2 To LastRow
        
    'Identify the tocker symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_Symbol = Cells(i, 1).Value
        Cells(Summary_Table, 9).Value = Ticker_Symbol
                
        'Identify the stock close price
        Price_Close = Cells(i, 6).Value
                
        'Calculate the difference between close price and open price
        Yearly_Change = Price_Close - Price_Open
                
        'Add stock change to summary table
        Cells(Summary_Table, 10).Value = Yearly_Change
                
            'Create Percent Change
            If (Price_Open = 0 And Price_Close = 0) Then
            Percent_Change = 0
                    
            ElseIf (Price_Open = 0 And Price_Close <> 0) Then
            Percent_Change = 1
                    
            Else
            Percent_Change = Yearly_Change / Price_Open
            Cells(Summary_Table, 11).Value = Percent_Change
            Cells(Summary_Table, 11).NumberFormat = "0.00%"
                    
                End If
                    
        'Calculate Total Volume
        Total_Volume = Total_Volume + Cells(i, 7).Value
                Cells(Summary_Table, 12).Value = Total_Volume
                
        'Add one to the summary table
                Summary_Table = Summary_Table + 1
                
        'reset the stock open price
        Price_Open = Cells(i + 1, 3)
                
                
        'Add cells with the same ticker
            Else
            Total_Volume = Total_Volume + Cells(i, 7).Value
                
            End If
            
        Next i
        
    ' Determine the Last Row of Yearly Change per WS
    C_LastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Add green and red to percent change column
    For j = 2 To C_LastRow
    
            'Create green color
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
                'Create red color
                ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 9
            
                End If
        Next j
        
    ' Add Greatest % Increase, % Decrease, and Total Volume to summary table
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
        
    'Loop through each rows to find the greatest value and its associate ticker
    For k = 2 To C_LastRow
            
        'Create the greatest percent increase
        If Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & C_LastRow)) Then
            Cells(2, 15).Value = Cells(k, 9).Value
            Cells(2, 16).Value = Cells(k, 11).Value
            Cells(2, 16).NumberFormat = "0.00%"
                
            'Create the greatest percent decrease
        ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & C_LastRow)) Then
            Cells(3, 15).Value = Cells(k, 9).Value
            Cells(3, 16).Value = Cells(k, 11).Value
            Cells(3, 16).NumberFormat = "0.00%"
                
        'Create the greatest total volume
        ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & C_LastRow)) Then
            Cells(4, 15).Value = Cells(k, 9).Value
            Cells(4, 16).Value = Cells(k, 12).Value
            
            End If
        
         Next k
        
    Next WS
        
End Sub
