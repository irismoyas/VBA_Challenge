Attribute VB_Name = "Module1"
Sub vbachallenge()
    
    For Each Ws In Worksheets
   

        Dim tickersymbol As String
        Dim tickervolume As Double
        tickervolume = 0
        Dim summarytickerrow As Integer
        summarytickerrow = 2
        
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double

       
        Ws.Cells(1, 9).Value = "Ticker"
        Ws.Cells(1, 10).Value = "Yearly Change"
        Ws.Cells(1, 11).Value = "Percent Change"
        Ws.Cells(1, 12).Value = "Total Stock Volume"

        
        lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        openprice = Ws.Cells(2, 3).Value

        For i = 2 To lastrow

        If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
        tickersymbol = Ws.Cells(i, 1).Value
        tickervolume = tickervolume + Ws.Cells(i, 7).Value
        Ws.Range("I" & summarytickerrow).Value = tickersymbol
        Ws.Range("L" & summarytickerrow).Value = tickervolume
        closeprice = Ws.Cells(i, 6).Value
        yearlychange = closeprice - openprice
        Ws.Range("J" & summarytickerrow).Value = yearlychange

         If openprice = 0 Then
         percentchange = 0
         Else
         percentchange = yearlychange / openprice
         End If
         
         Ws.Range("K" & summarytickerrow).Value = percentchange
         Ws.Range("K" & summarytickerrow).NumberFormat = "0.00%"
         
         summarytickerrow = summarytickerrow + 1
         tickervolume = 0
         openprice = Ws.Cells(i + 1, 3)
         
         Else
         
         tickervolume = tickervolume + Ws.Cells(i, 7).Value
         End If
         Next i
         
         lastrowsummarytable = Ws.Cells(Rows.Count, 9).End(xlUp).Row
       
       For i = 2 To lastrowsummarytable
            If Ws.Cells(i, 10).Value > 0 Then
                Ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                Ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

        Ws.Cells(2, 15).Value = "Greatest % Increase"
        Ws.Cells(3, 15).Value = "Greatest % Decrease"
        Ws.Cells(4, 15).Value = "Greatest Total Volume"
        Ws.Cells(1, 16).Value = "Ticker"
        Ws.Cells(1, 17).Value = "Value"

    
        For i = 2 To lastrowsummarytable
            
            If Ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(Ws.Range("K2:K" & lastrowsummarytable)) Then
               Ws.Cells(2, 16).Value = Ws.Cells(i, 9).Value
                Ws.Cells(2, 17).Value = Ws.Cells(i, 11).Value
                Ws.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf Ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(Ws.Range("K2:K" & lastrowsummarytable)) Then
                Ws.Cells(3, 16).Value = Ws.Cells(i, 9).Value
                Ws.Cells(3, 17).Value = Ws.Cells(i, 11).Value
                Ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf Ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(Ws.Range("L2:L" & lastrowsummarytable)) Then
                Ws.Cells(4, 16).Value = Ws.Cells(i, 9).Value
                Ws.Cells(4, 17).Value = Ws.Cells(i, 12).Value
            
            End If
        
        Next i
        
        Ws.Columns("I:L").EntireColumn.AutoFit
        Ws.Columns("O:Q").EntireColumn.AutoFit

        Next Ws
        
End Sub

