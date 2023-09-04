Attribute VB_Name = "Module1"
Sub Yearlystock()

 Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim YearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    
    For Each ws In ThisWorkbook.Worksheets
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "YearlyChange"
    ws.Cells(1, 11).Value = "PercentageChange"
    ws.Cells(1, 12).Value = "TotalStockVolume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest TotalVolume"
    
    Dim summarytable As Long
    summarytable = 2


    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Row = 2
    'firstRow = 2
    Year_open = ws.Cells(2, 3).Value
     For i = 2 To lastRow
   
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Collect ticker symbol
        'ws.Cells(Row, 9).Value = ws.Cells(i, 1).Value
        ticker = ws.Cells(i, 1).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        ws.Range("I" & summarytable).Value = ticker
        ws.Range("L" & summarytable).Value = totalVolume
        'Put in yearly change
        Year_close = ws.Cells(i, 6).Value
        YearlyChange = Year_close - Year_open
        'Colour in yearly change
        If ws.Cells(Row, 10).Value < 0 Then
            ws.Cells(Row, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(Row, 10).Interior.ColorIndex = 4
        End If
        
        percentChange = (YearlyChange / Year_open)
         'convert decimal value to percentage
        
        ws.Range("K" & summarytable).NumberFormat = "0.00%"
        ws.Range("K" & summarytable).Value = percentChange
        
        'Put in percent change
      
       Year_open = ws.Cells(i + 1, 3).Value
        ws.Range("J" & summarytable).Value = YearlyChange
        
        
        'Update the firstRow counter
       summarytable = summarytable + 1
      Else
        'Reset stock vol
       totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        
        
    
    End If
    
  
  Next i
  
  Next ws
        
End Sub

