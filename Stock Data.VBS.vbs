FINAL



Sub Stocktable()
    Dim volume As Double
    Dim ticker As String
    Dim stockopen As Double
    Dim stockclose As Double
    Dim stockchange As Double
    Dim percentchange As Double
    Dim start As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Stocktable As Integer
    For Each ws In Worksheets
        j = 0
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        volume = 0
        Stocktable = 2
        start = 2
        
        'Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox LastRow
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Cells(Stocktable, 9).Value = ticker
                volume = volume + ws.Cells(i, 7).Value
                ws.Cells(Stocktable, 12).Value = volume
                'annual change
                stockopen = ws.Cells(start, 3).Value
                stockclose = ws.Cells(i, 6).Value
                stockchange = stockclose - stockopen
                ws.Cells(Stocktable, 10).Value = stockchange
                If ws.Cells(Stocktable, 10).Value > 0 Then
                    ws.Cells(Stocktable, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Stocktable, 10).Interior.ColorIndex = 3
                End If
                'Percent Change
                If stockopen = 0 Then
                    percentchange = 0
                Else
                    percentchange = stockchange / stockopen
                End If
                ws.Cells(Stocktable, 11).Value = percentchange
                ws.Cells(Stocktable, 11).NumberFormat = "0.00%"
                
                
                
                Stocktable = Stocktable + 1
                start = i + 1
                volume = 0
                stockchange = 0
                percentchange = 0
            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
         Next i
    Next ws
End Sub
