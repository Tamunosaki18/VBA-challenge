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
    Dim vol As Long
    Dim max_per_inc As Double
    Dim min_per_inc As Double
    Dim gr_total As Double
    
    For Each ws In ThisWorkbook.Worksheets
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "YearlyChange"
    ws.Cells(1, 11).Value = "PercentageChange"
    ws.Cells(1, 12).Value = "TotalStockVolume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest TotalVolume"
    
    Dim summarytable As Long
    summarytable = 2
    YearlyChange = 0
    percentChange = 0
    totalVolume = 0
    Year_open = 0


    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
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
        totalVolume = 0
        'Put in yearly change
        Year_close = ws.Cells(i, 6).Value

        YearlyChange = Year_close - Year_open
        'Color coding negative and positive numbers in the YearlyChange column
        
        If YearlyChange >= 0 Then
            ws.Range("J" & summarytable).Interior.ColorIndex = 4 ' Green color
        Else
            ws.Range("J" & summarytable).Interior.ColorIndex = 3 ' Red color
            
        End If
            
        'percentChange Calculation
        ' handle division by zero error
        If Year_open = 0 Then
            percentChange = 0
        Else
            percentChange = (YearlyChange / Year_open)
        End If
        
        'Update Year_open
        Year_open = ws.Cells(i + 1, 3).Value
         
         'convert decimal value to percentage
        
        ws.Range("K" & summarytable).NumberFormat = "0.00%"
        ws.Range("K" & summarytable).Value = percentChange
        
        'Pull opening value of next ticker
      Year_open = ws.Cells(i + 1, 3).Value
        
        'Print the yearly change in column J
        ws.Range("J" & summarytable).Value = YearlyChange
   ' If totalVolume > maxTotalVolume Then
                    'maxTotalVolume = totalVolume
                   ' tickerMaxTotalVolume = ticker
       
       'Add one to the summary table
        summarytable = summarytable + 1
      Else
        'Add to the Total Stock Volume
      totalVolume = totalVolume + ws.Cells(i, 7).Value
        
    End If
    


    
    
  Next i
    
   Next ws
        
End Sub

