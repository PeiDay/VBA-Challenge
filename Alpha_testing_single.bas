Attribute VB_Name = "Module2"

Sub alphabetical_testing_single_ws()

Dim iTickerName, iMaxPerTickerName, iMinPerTickerName, iMaxVolTickerName As String
Dim iTickerVol, iMaxVolTicker, iMaxVolTickerRow As Double
Dim YOpen, YClose, NextYOpen, YChange, YPerChange As Double
Dim i, LastRow, SumRow, LastSumRow As Double
Dim iMaxPerTicker, iMinPerTicker, iMaxPerTickerRow, iMinPerTickerRow As Double

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'setting up the summary tables
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"

'    Cells(1, 14).Value = "YOpen"     'make to comment after testing
'    Cells(1, 15).Value = "YClose"    'make to comment after testing
    
    iTickerVol = 0
    YOpen = 0
    SumRow = 1

    'calculate for each Ticker
    For i = 2 To LastRow
    
        'get the open price for the first ticker
        If YOpen = 0 Then
            YOpen = Cells(i, 3).Value
        End If
        
        iTickerVol = iTickerVol + Cells(i, 7).Value
            
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            iTickerName = Cells(i, 1).Value
            Cells(SumRow + 1, 10).Value = iTickerName
            
            YClose = Cells(i, 6).Value
'            Cells(SumRow + 1, 15).Value = YClose 'make to comment after testing
'            Cells(SumRow + 1, 14).Value = YOpen 'make to comment after testing
           
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

            Cells(SumRow + 1, 11).Value = YChange
        
            'conditional formatting for Yearly change
            If YChange > 0 Then
                Cells(SumRow + 1, 11).Interior.ColorIndex = 4
                ElseIf YChange < 0 Then
                Cells(SumRow + 1, 11).Interior.ColorIndex = 3
            End If

            YPerChange = CDbl(YChange / YOpen)
            Cells(SumRow + 1, 12).Value = YPerChange
            Cells(SumRow + 1, 12).NumberFormat = "0.00%"
            Cells(SumRow + 1, 13).Value = iTickerVol
            
            'reset: volume, open price, and summary table row
            iTickerVol = 0
            SumRow = SumRow + 1
            
            'prepare for next open price
            YOpen = Cells(i + 1, 3).Value
        End If
        
    Next i
   

    'second summary table for Greatest % and total
    LastSumRow = Cells(Rows.Count, 10).End(xlUp).Row

    Cells(2, 16).Value = "Greatest % Increase"
    Cells(3, 16).Value = "Greatest % Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"
    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
        
    'Get the greatest % Increase and Decrease
    Set PerRng = Range("L:L")

    iMaxPerTicker = WorksheetFunction.Max(PerRng)
    iMaxPerTickerRow = WorksheetFunction.Match(iMaxPerTicker, PerRng, 0) + PerRng.Row - 1
    iMaxPerTickerName = Cells(iMaxPerTickerRow, 10).Value
    Cells(2, 17).Value = iMaxPerTickerName
    Cells(2, 18).Value = iMaxPerTicker
    Cells(2, 18).NumberFormat = "0.00%"

    iMinPerTicker = WorksheetFunction.Min(PerRng)
    iMinPerTickerRow = WorksheetFunction.Match(iMinPerTicker, PerRng, 0) + PerRng.Row - 1
    iMinPerTickerName = Cells(iMinPerTickerRow, 10).Value
    Cells(3, 17).Value = iMinPerTickerName
    Cells(3, 18).Value = iMinPerTicker
    Cells(3, 18).NumberFormat = "0.00%"
    
    'get the highest Volume
    Set VolRng = Range("M:M")
    
    iMaxVolTicker = WorksheetFunction.Max(VolRng)
    iMaxVolTickerRow = WorksheetFunction.Match(iMaxVolTicker, VolRng, 0) + VolRng.Row - 1
    iMaxVolTickerName = Cells(iMaxVolTickerRow, 10).Value
    Cells(4, 17).Value = iMaxVolTickerName
    Cells(4, 18).Value = iMaxVolTicker
    
    'format for display
    Columns("A:R").AutoFit
    

End Sub

