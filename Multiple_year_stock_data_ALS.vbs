Attribute VB_Name = "Module1"
Sub extractstockvolume()

For Each ws In Worksheets

        
        volumetotal = 0
        TableRow = 2
        

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'For i = 2 To LastRow
        
        For i = 2 To LastRow
        

        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Get ticker symbol
        
        
        Ticker = ws.Cells(i, 1).Value
        'priceopen = ws.Cells(i, 3).Value
        
        
        'Tally volume total
        
        volumetotal = volumetotal + ws.Cells(i, 7).Value
        
        'Print the ticker symbol and volume totals
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        ws.Range("I" & TableRow).Value = Ticker
        ws.Range("J" & TableRow).Value = volumetotal
        
        TableRow = TableRow + 1
        
        volumetotal = 0
        
        Else
        
        volumetotal = volumetotal + ws.Cells(i, 7).Value
        End If
        Next i
        Next ws
        
End Sub


 

