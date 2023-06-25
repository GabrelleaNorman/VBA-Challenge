Attribute VB_Name = "Module11"
Sub StockDataTicker()

   
   For Each ws In Worksheets
    
    Dim ThisWorkbook As String
    'Current row
        Dim i As Double
    'Starting row for ticker
        Dim j As Double
    'Index counter to fill Ticker row
        Dim TickerCount As Long
    'Ticker Symbol Column A Last Row
        Dim TickerSymbolA As Long
    'Ticker Symbol Column I Last Row
        Dim TickerSymbolI As Long
    'Variable for percent change
        Dim PercentChange As Double
       
        
    'Name the Worksheets
        ThisWorkbook = ws.Activate
        
    'Input Column headers for Ticker, Yearly Change, Percent Change, and Total Stock Volume (Columns I, J, K, L)
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
      
    
        
    'Set Ticker Counter to first row
        TickerCount = 2
        
    'Set start row to 2
        j = 2
        
    'Locate last data cell in Ticker Column (A)
        TickerSymbolA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'MsgBox ("Last row in column A is " & TickerSymbolA)
        
    
    'Loop through all rows
            For i = 2 To TickerSymbolA
            
            'Identify Change in the ticker symbol
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
            'Add Ticker Symbol to Column
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                
            'Calculate the values for Yearly Change (Closing Price - Open Price)
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
            'Add Conditional formating to display change
                    If ws.Cells(TickerCount, 10).Value < 0 Then
                
            'Change cell color to red for Negative Change
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
        Else
                
            'Change cell color to green for Positive Change
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
        End If
                    
            'Calculate the percent change from opening price to the closing price(Yearly Change/Opening Price)
                If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
            'Percent formating
                    ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                    
        Else
                
                    ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
        End If
                    
            'Calculate the Total Stock Volume of Stock
                    ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickerCount by 1
                TickerCount = TickerCount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
        Next i
            
        'Create Column Header
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Variable for greatest percentage increase
        Dim GreatestIncrease As Double
        'Variable for greatest percentage decrease
        Dim GreatestDecrease As Double
        'Variable for greatest total volume
        Dim GreatestTotVolume As Double
        
        
        
        
        'Find last non-blank cell in Ticker Symbol Column I
        TickerSymbolI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & TickerSymbolI)
        
        'Prepare for summary
        GreatestTotVolume = ws.Cells(2, 12).Value
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To TickerSymbolI
            
                'Calculate greatest total volume
                If ws.Cells(i, 12).Value > GreatestTotVolume Then
                GreatestTotVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestTotVolume = GreatestTotVolume
                
                End If
                
                'Calculate greatest percent increase
                If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestIncrease = GreatestIncrease
                
                End If
                
                'Calculate greatest percent decrease
                If ws.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestDecrease = GreatestDecrease
                
                End If
                
           'Calculate value for Greatest%Increase,Greatest%Decrease,and Greatest Total Volume in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestTotVolume, "Scientific")
            
            Next i
            
        
            
    Next ws
        
End Sub
