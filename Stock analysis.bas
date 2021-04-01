Attribute VB_Name = "Module1"
Sub StockAnalysis()

    
    For Each ws In Worksheets

        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalStockVolume As Double
        Dim SummaryTable As Long
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        Dim PercentChange As Double
        TotalStockVolume = 0
        SummaryTable = 2
        PreviousAmount = 2

        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

           
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                TickerName = ws.Cells(i, 1).Value
              
                ws.Range("I" & SummaryTable).Value = TickerName
                ws.Range("L" & SummaryTable).Value = TotalStockVolume
                
                TotalTickerVolume = 0


                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTable).Value = YearlyChange

                
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                
                ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTable).Value = PercentChange

                
                If ws.Range("J" & SummaryTable).Value >= 0 Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                End If
            
               
                SummaryTable = SummaryTable + 1
                PreviousAmount = i + 1
                End If
            Next i
        
        ws.Columns("I:L").AutoFit

    Next ws

End Sub


