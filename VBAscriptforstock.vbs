Sub Multi_Year_StockData()

For Each ws In Worksheets

    Dim i, j As Integer
    Dim openingPrice, closingPrice, yearlyChange As Double
    Dim tickerCount As Integer
    Dim percentChange, totalStock, greatIncrease, greatDecrease, greatStock As Double
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    'part = 1
    lastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    j = 2
    
    For i = 2 To lastRowA
        
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            openingPrice = ws.Cells(i, 3).Value
            totalStock = 0
       End If
       
       totalStock = totalStock + ws.Cells(i, 7).Value
       
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            closingPrice = ws.Cells(i, 6).Value
            ticker = ws.Cells(i, 1).Value
            yearlyChange = closingPrice - openingPrice
            percentChange = yearlyChange / openingPrice
        
            
            ws.Cells(j, 9).Value = ticker
            ws.Cells(j, 10).Value = yearlyChange
            ws.Cells(j, 11).Value = percentChange
            ws.Cells(j, 12).Value = totalStock
            
            'color condition
            If ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 4
            End If
            
            'percent formating
            If ws.Cells(j, 3).Value <> 0 Then
            ws.Cells(j, 11).Value = Format(percentChange, "Percent")
            Else
            ws.Cells(j, 11).Value = Format(0, "Percent")
            End If
            
           j = j + 1
    
        End If
        
    Next i
    
    'part = 2
    'summary report
    lastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    greatIncrease = ws.Cells(2, 11).Value
    greatDecrease = ws.Cells(2, 11).Value
    greatStock = ws.Cells(2, 12).Value
    
    For i = 2 To lastRowI
    
        'find greatest increase
        If ws.Cells(i, 11).Value > greatIncrease Then
            greatIncrease = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        Else
            greatIncrease = greatIncrease
        End If
        
        'find greatest decrease
        If ws.Cells(i, 11).Value < greatDecrease Then
            greatDecrease = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        Else
            greatDecrease = greatDecrease
        End If
        
        'find greatest stock volume
        If ws.Cells(i, 12).Value > greatStock Then
            greatStock = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        Else
            greatStock = greatStock
        End If
        
    Next i
    
    ws.Cells(2, 17).Value = Format(greatIncrease, "Percent")
    ws.Cells(3, 17).Value = Format(greatDecrease, "Percent")
    ws.Cells(4, 17).Value = Format(greatStock, "Scientific")
    
Next ws
    
End Sub


