Attribute VB_Name = "Module1"

Sub StockMarketMultiYear()

Dim iTickerName, iMaxPerTickerName, iMinPerTickerName, iMaxVolTickerName As String
Dim iTickerVol, iMaxVolTicker, iMaxVolTickerRow As Double
Dim YOpen, YClose, NextYOpen, YChange, YPerChange As Double
Dim i, LastRow, SumRow, LastSumRow As Double
Dim iMaxPerTicker, iMinPerTicker, iMaxPerTickerRow, iMinPerTickerRow As Double

For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'setting up the summary tables
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"

'    ws.Cells(1, 14).Value = "YOpen"     'make to comment after testing
'    ws.Cells(1, 15).Value = "YClose"    'make to comment after testing
    
    iTickerVol = 0
    YOpen = 0
    SumRow = 1

    'calculate for each Ticker
    For i = 2 To LastRow
    
        'get the open price for the first ticker
        If YOpen = 0 Then
            YOpen = ws.Cells(i, 3).Value
        End If
        
        iTickerVol = iTickerVol + ws.Cells(i, 7).Value
            
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            iTickerName = ws.Cells(i, 1).Value
            ws.Cells(SumRow + 1, 10).Value = iTickerName
            
            YClose = ws.Cells(i, 6).Value
'            ws.Cells(SumRow + 1, 15).Value = YClose 'make to comment after testing
'            ws.Cells(SumRow + 1, 14).Value = YOpen 'make to comment after testing
            
            'What if YOpen/Yclose =0
            If YOpen = 0 And YClose = 0 Then
                YChange = 0
                YPerChange = 0
                
                ElseIf YOpen = 0 And YClose > 0 Then
                    YChange = CDbl(YClose - YOpen)
                    YPerChange = 1
                 
                Else
                    YChange = YClose - YOpen
                    YPerChange = YChange / YOpen
            End If

            ws.Cells(SumRow + 1, 11).Value = YChange
        
            'conditional formatting for Yearly change
            If YChange > 0 Then
                ws.Cells(SumRow + 1, 11).Interior.ColorIndex = 4
                ElseIf YChange < 0 Then
                ws.Cells(SumRow + 1, 11).Interior.ColorIndex = 3
            End If

            ws.Cells(SumRow + 1, 12).Value = YPerChange
            ws.Cells(SumRow + 1, 12).NumberFormat = "0.00%"
            ws.Cells(SumRow + 1, 13).Value = iTickerVol
            
            'reset: volume, open price, and summary table row
            iTickerVol = 0
            SumRow = SumRow + 1
            
            'prepare for next open price
            YOpen = ws.Cells(i + 1, 3).Value
        End If
        
    Next i
   

    'second summary table for Greatest % and total
    LastSumRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
        
    'Get the greatest % Increase and Decrease
    Set PerRng = ws.Range("L:L")

    iMaxPerTicker = WorksheetFunction.Max(PerRng)
    iMaxPerTickerRow = WorksheetFunction.Match(iMaxPerTicker, PerRng, 0) + PerRng.Row - 1
    iMaxPerTickerName = ws.Cells(iMaxPerTickerRow, 10).Value
    ws.Cells(2, 17).Value = iMaxPerTickerName
    ws.Cells(2, 18).Value = iMaxPerTicker
    ws.Cells(2, 18).NumberFormat = "0.00%"

    iMinPerTicker = WorksheetFunction.Min(PerRng)
    iMinPerTickerRow = WorksheetFunction.Match(iMinPerTicker, PerRng, 0) + PerRng.Row - 1
    iMinPerTickerName = ws.Cells(iMinPerTickerRow, 10).Value
    ws.Cells(3, 17).Value = iMinPerTickerName
    ws.Cells(3, 18).Value = iMinPerTicker
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
    'get the highest Volume
    Set VolRng = ws.Range("M:M")
    
    iMaxVolTicker = WorksheetFunction.Max(VolRng)
    iMaxVolTickerRow = WorksheetFunction.Match(iMaxVolTicker, VolRng, 0) + VolRng.Row - 1
    iMaxVolTickerName = ws.Cells(iMaxVolTickerRow, 10).Value
    ws.Cells(4, 17).Value = iMaxVolTickerName
    ws.Cells(4, 18).Value = iMaxVolTicker
    
    'format for display
    ws.Columns("A:R").AutoFit
    
Next ws

End Sub

