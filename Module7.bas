Attribute VB_Name = "Module7"
'Tickers
Sub creditcard()

For Each ws In ThisWorkbook.Worksheets

    Dim i As Long
    Dim openingprice As Double
    Dim closingprice As Double
    Dim yearlychange As Double
    Dim totalvolume As Double
    Dim percentchange As Double
    Dim tickertime As Integer
    Dim SmallTicker As String
    Dim LargeTicker As String
    Dim maxincrease As Double
    Dim maxdecrease As Double
    Dim largeTickerValue As String
    Dim volumeticker As String
    Dim MAXVOLUME As Double
    
    maxincrease = 0
    maxdecrease = 0
    MAXVOLUME = 0
    
    
    tickertime = 2
    
 

    For i = 2 To 753001

            ' Stock Volume
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        
        ' find last row of Ticker
        
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            ws.Cells(tickertime, 9) = ws.Cells(i, 1)
            
            
            
            ' closeing price
        closingprice = ws.Cells(i, 6).Value
            
        yearlychange = closingprice - openingprice
            ws.Cells(tickertime, 10).Value = yearlychange
            If yearlychange > 0 Then
                'green
                    ws.Cells(tickertime, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlychange <= 0 Then
                'red
                    ws.Cells(tickertime, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
            
            'percentChange
            If openingprice <> 0 Then
            
            percentchange = (yearlychange / openingprice)
            ws.Cells(tickertime, 11).Value = percentchange
            ws.Cells(tickertime, 11).NumberFormat = "0.00%"
            Else
            ws.Cells(tickertime, 11).Value = 0
            End If
            
             ' Output total stock volume
            ws.Cells(tickertime, 12).Value = totalvolume
            
            ' Increment tickerTime
            tickertime = tickertime + 1
            
            ' Reset totalvolume for the next ticker
            totalvolume = 0
            
        'findsFirstRow
        
               
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
' storeOPEN
            openingprice = ws.Cells(i, 3)
            
            
            End If
            'GreatestIncrease
         If ws.Cells(tickertime, 11).Value > maxincrease Then
        maxincrease = ws.Cells(tickertime, 11).Value
        LargeTicker = ws.Cells(tickertime, 9).Value
         ws.Cells(2, 16).NumberFormat = "0.00%"
        
   'GreatestDecrease
        
        End If
        If ws.Cells(tickertime, 11).Value < maxdecrease Then
        maxdecrease = ws.Cells(tickertime, 11).Value
        SmallTicker = ws.Cells(tickertime, 9).Value
        ws.Cells(3, 16).NumberFormat = "0.00%"
        End If
        
    'GreatTotalVolume
        If ws.Cells(tickertime, 12) > MAXVOLUME Then
        MAXVOLUME = ws.Cells(tickertime, 12).Value
        volumeticker = ws.Cells(tickertime, 9).Value
        End If
        
        
        

    
           
        
    Next i

    ' name
    ws.Cells(4, 15).Value = volumeticker
    ws.Cells(4, 16).Value = MAXVOLUME
    ws.Cells(3, 15).Value = SmallTicker
    ws.Cells(3, 16).Value = maxdecrease
    ws.Cells(2, 16).Value = maxincrease
    ws.Cells(2, 15).Value = LargeTicker
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 14) = "Summary Stats"
    ws.Cells(1, 15) = "Tickers"
    ws.Cells(1, 16) = "Value"
    ws.Cells(2, 14) = "Greatest % increase"
    ws.Cells(3, 14) = "Greatest % decrease"
    ws.Cells(4, 14) = "Greatest Total Volume"
    
    
Next ws

End Sub

