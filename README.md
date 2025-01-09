The only thing that I used outside the class lecture was how to find the last row

lastRow = Cells(Rows.Count, "A").End(xlUp).Row

https://stackoverflow.com/questions/38882321/better-way-to-find-last-used-row

Following is my script:

Sub StockMarket()
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim tickerMaxIncrease As String
    Dim tickerMaxDecrease As String
    Dim tickerMaxVolume As String
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

    openPrice = Cells(2, "C").Value
    
    totalVolume = 0
   
    outputRow = 2

    For i = 2 To lastRow
       
        totalVolume = totalVolume + Cells(i, "G").Value
        
        ' Check if the ticker changes or it's the last row
      
        If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
            
            ticker = Cells(i, "A").Value
            closePrice = Cells(i, "F").Value
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
            percentChange = (quarterlyChange / openPrice) * 100
            Else
            percentChange = 0
            End If
                                                         
            
            Cells(outputRow, "I").Value = ticker
            Cells(outputRow, "J").Value = quarterlyChange
            Cells(outputRow, "K").Value = percentChange
            Cells(outputRow, "L").Value = totalVolume

outputRow = outputRow + 1
           
            openPrice = Cells(i + 1, "C").Value
           
            totalVolume = 0
     
        End If
   
    Next i

'Greatest % increase, Greatest % decrease, Greatest total volume

    For i = 2 To lastRow
       
        If Cells(i, "K").Value > maxIncrease Then
          
            maxIncrease = Cells(i, "K").Value
           
            tickerMaxIncrease = Cells(i, "I").Value
     
        End If

        If Cells(i, "K").Value < maxDecrease Then
           
            maxDecrease = Cells(i, "K").Value
            
            tickerMaxDecrease = Cells(i, "I").Value
       
        End If

        If Cells(i, "L").Value > maxVolume Then
           
            maxVolume = Cells(i, "L").Value
            
            tickerMaxVolume = Cells(i, "I").Value
       
        End If
   
    Next i

        
    ' Output max values
    Cells(1, "I") = "Ticker"
    Cells(1, "J") = "Quartely Charge"
    Cells(1, "K") = "Percent Charge"
    Cells(1, "L") = "Total Stock Valume"
    Cells(1, "P") = "Ticker"
    Cells(1, "Q") = "Value"
    Cells(2, "P").Value = tickerMaxIncrease
    Cells(3, "P").Value = tickerMaxDecrease
    Cells(4, "P").Value = tickerMaxVolume
    Cells(2, "Q").Value = maxIncrease
    Cells(3, "Q").Value = maxDecrease
    Cells(4, "Q").Value = maxVolume
    Cells(2, "O").Value = "Greatest % increase"
    Cells(3, "O").Value = "Greatest % decrease"
    Cells(4, "O").Value = "Greatest total volume"

    ' Apply conditional formatting for Quarterly Change column
    
    For i = 2 To lastRow
    
    If (Cells(i, 10).Value < 0) Then
    
    Cells(i, 10).Interior.ColorIndex = 3
    
    ElseIf (Cells(i, 10).Value > 0) Then

    Cells(i, 10).Interior.ColorIndex = 4
    
    End If
    
Next i

    
End Sub


