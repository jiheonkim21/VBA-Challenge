Sub stockChanges()

' Script  will loop through all the stocks for one year and output the following information.
'  * The ticker symbol.
'  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'  * The total stock volume of the stock.

    Dim ws As Worksheet

    
    For Each ws In Worksheets
        ws.Activate
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
    
        Dim startPrice As Double
        Dim endPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalStockVolume As Double
        
        Dim greatestPercIncrease As Double
        Dim greatestPercDecrease As Double
        Dim greatestTotalVolume As Double
        
        greatestPercIncrease = 0
        greatestPercDecrease = 0
        greatestTotalVolume = 0
        
        
        startPrice = ws.Range("F2").Value
        endPrice = ws.Range("F2").Value
        yearlyChange = endPrice - startPrice
        percentageChange = yearlyChange / startPrice
        totalStockVolume = 0
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        
        For i = 2 To lastRowState
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                endPrice = ws.Cells(i, 6).Value
                yearlyChange = (endPrice - startPrice)
                
                If totalStockVolume > greatestTotalVolume Then
                    ws.Range("P4") = ws.Cells(i, 1).Value
                    ws.Range("Q4") = totalStockVolume
                    greatestTotalVolume = totalStockVolume
                End If
                
                If yearlyChange = 0 Then
                    percentageChange = 0
                ElseIf startPrice = 0 Then
                    percentageChange = 0
                Else
                    percentageChange = yearlyChange / startPrice
                End If
                
                If percentageChange < greatestPercDecrease Then
                    ws.Range("P3") = ws.Cells(i, 1).Value
                    ws.Range("Q3") = percentageChange
                    greatestPercDecrease = percentageChange
                End If
                
                If percentageChange > greatestPercIncrease Then
                    ws.Range("P2") = ws.Cells(i, 1).Value
                    ws.Range("Q2") = percentageChange
                    greatestPercIncrease = percentageChange
                End If
                
                
                startPrice = ws.Cells(i + 1, 5).Value
                ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Summary_Table_Row, 10).Value = yearlyChange
                If (yearlyChange < 0) Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                ElseIf (yearlyChange > 0) Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(Summary_Table_Row, 11).Value = percentageChange
                ws.Cells(Summary_Table_Row, 12).Value = totalStockVolume
                Summary_Table_Row = Summary_Table_Row + 1
                totalStockVolume = 0
            End If
                
        Next i
        ws.Columns("I:Q").AutoFit
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next ws
    
End Sub