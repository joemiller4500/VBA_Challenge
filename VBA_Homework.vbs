
Sub Stocks()
    
    'Declare Variables
    Dim columnLength As Double
    Dim lastTick As String
    Dim opening As Double
    Dim closing As Double
    Dim tickCount As Double
    Dim stockTot As Double
    Dim yearChange As Double
    Dim percChange As Double
    Dim temp As Double
    Dim grPerInc As Double
    Dim grInc As String
    Dim grPerDec As Double
    Dim grDec As String
    Dim grTotVol As Double
    Dim grTot As String
    
    'Title Columns
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Cells(2, 15) = "Greatest Percent Increase"
    Cells(3, 15) = "Greatest Percent Decrease"
    Cells(4, 15) = "Greatest Total Stock Volume"
    
    'Initialize Variables
    columnLength = Cells(Rows.Count, 1).End(xlUp).Row
    lastTick = Cells(2, 1)
    opening = Cells(2, 3)
    tickCount = 2
    
    'Iterate through rows
    For i = 2 To columnLength
        
        'If ticker hasn't changed
        If Cells(i, 1) = lastTick Then
            stockTot = stockTot + Cells(i, 7)
            
        'If ticker has changed
        Else
            temp = i - 1
            closing = Cells(temp, 6)
            yearChange = (opening - closing)
            percChange = yearChange / (opening + 1E-08)
            Cells(tickCount, 9).Value = lastTick
            Cells(tickCount, 10).Value = yearChange
            
            If yearChange > 0 Then
                Cells(tickCount, 10).Interior.ColorIndex = 4
                
            Else
                Cells(tickCount, 10).Interior.ColorIndex = 3
                End If
                
            Cells(tickCount, 11).Value = percChange
            Cells(tickCount, 12).Value = stockTot
            opening = Cells(i, 3)
            tickCount = tickCount + 1
            stockTot = 0
            End If
            
            'Bonus
            If percChange > grPerInc Then
                grPerInc = percChange
                grInc = lastTick
                End If
                
            If percChange < grPerDec Then
                grPerDec = percChange
                grDec = lastTick
                End If
                
            If stockTot > grTotVol Then
                grTotVol = stockTot
                grTot = lastTick
                End If
                
            lastTick = Cells(i, 1)
            
        Next i
            
 'Populate bonus
 Cells(2, 16) = grInc
 Cells(2, 17) = grPerInc
 Cells(3, 16) = grDec
 Cells(3, 17) = grPerDec
 Cells(4, 16) = grTot
 Cells(4, 17) = grTotVol
 
 Columns(11).NumberFormat = "0.00%"
 Cells(2, 17).NumberFormat = "0.00%"
 Cells(3, 17).NumberFormat = "0.00%"
            

End Sub




Type your solution here