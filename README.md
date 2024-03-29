Sub Ticker()


'TICKER SYMBOL
        Dim ws As Worksheet
        Dim lastRow As Long
        For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row in column A
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

          'Set the headers for column I-L
          ws.Cells(1, "I").Value = "Ticker"
          ws.Cells(1, "J").Value = "Yearly Change"
          ws.Cells(1, "K").Value = "Percent Change"
          ws.Cells(1, "L").Value = "Total Stock Volume"

         ' Loop through each row in column A and copy the value to column I
        For i = 2 To lastRow
            ws.Cells(i, "I").Value = ws.Cells(i, "A").Value
            'remove repeating ticker symbols
    
        Next i
        ws.Range("I1:I" & lastRow).RemoveDuplicates Columns:=1, Header:=xlNoY
    'YEARLY CHANGE,PERCENT CHANGE
        'open & close price, yearly change, volume count
        Dim close_price, open_price, yearly_change, volume_count As Double
        Dim current_tick As Integer
        
        
        'assign indexes for column I
        current_tick = 2
    
        volume_count = 0
        'first open price value
        open_price = ws.Cells(2, 3).Value

        'loop through each row in colun A and determine if it's a different ticker
        For i = 2 To lastRow
        
            volume_count = volume_count + Cells(i, 7)
         'Check to see if the ticker symbol is different to get the close value “F”
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                close_price = ws.Cells(i, 6).Value
                'calculate yearly change and assign to tick and percent change
                ws.Cells(current_tick, 10).Value = close_price - open_price
                ws.Cells(current_tick, 11).Value = ((close_price - open_price) / (open_price)) * 100
                'assign total volume to column L
                ws.Cells(current_tick, 12).Value = volume_count
                'reset volume
                volume_count = 0
                 'COLOR the percent change column
                    If ws.Cells(current_tick, 10).Value < 0 Then
                        'red if less than 0
                        ws.Cells(current_tick, 10).Interior.ColorIndex = 3
                        'green if greater than 0
                    Else
                        ws.Cells(current_tick, 10).Interior.ColorIndex = 4
                    End If
                'move to next ticker
                current_tick = current_tick + 1
                open_price = ws.Cells(i + 1, 3).Value
            End If
        Next i
        
        
        
        'NEW TABLE
        'make new table
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest percent increase"
        ws.Cells(3, "O").Value = "Greatest percent decrease"
        ws.Cells(4, "O").Value = "Greatest total volume"
        
        
        'initate new variables: percent increase, percent decrease, and greatest total volume
        Dim maxValue, minValue, maxTV As Double
        Dim maxVtick, minVtick, maxTVtick As Double
        
        'find greatest percent increase
        maxValue = Application.Max(ws.Range("K:K"))
        ws.Cells(2, "Q").Value = maxValue
        'find the ticker symbol that corresponds to the maxValue
            maxVtick = Application.Match(maxValue, ws.Range("K:K"), 0)
            ws.Cells(2, "P").Value = ws.Cells(maxVtick, "I")
            
        
        'Cells(2,"P").Value=TICKER SYMBOL
         'find greatest percent decrease
        minValue = Application.Min(ws.Range("K:K"))
        ws.Cells(3, "Q").Value = minValue
        'find the ticker symbol that corresponds to the minValue
            minVtick = Application.Match(minValue, ws.Range("K:K"), 0)
            ws.Cells(3, "P").Value = ws.Cells(minVtick, "I")
        
        'find greatest total volume
        maxTV = Application.Max(ws.Range("L:L"))
        ws.Cells(4, "Q").Value = maxTV
         'find the ticker symbol that corresponds to the maxTVtick
            maxTVtick = Application.Match(maxTV, ws.Range("L:L"), 0)
            ws.Cells(4, "P").Value = ws.Cells(maxTVtick, "I")
      
      Next ws
End Sub




