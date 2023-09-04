Attribute VB_Name = "Module2"
Sub summarytable_2()

Dim summarytabe

  For Each ws In ThisWorkbook.Worksheets

lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    max_per_inc = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))


'Identify the location of the matching ticker
ticker_one = WorksheetFunction.Match(max_per_inc, ws.Range("K2:K" & lastRow), 0)
'Finding the matching ticker to the location
ticker_one = ticker_one + 1
maxPosition = ws.Cells(ticker_one, 9).Value

'Print maxposition into the corresponding cell
ws.Cells(2, 16).Value = maxPosition

min_per_inc = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))

ticker_two = WorksheetFunction.Match(min_per_inc, ws.Range("K2:K" & lastRow), 0)

'Finding the matching ticker to the location
    ticker_two = ticker_two + 1
    minPosition = ws.Cells(ticker_two, 9).Value

'Print minPosition into the corresponding cell
ws.Cells(3, 16).Value = minPosition

'Formatting greatest increase and decrease to percentage
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

'Identifying greatest totalvolume for the second summarytable
gr_vol = WorksheetFuntion.Max(ws.Range("L2:L" & lastRow))
'Identify location of the matching ticker
ticker_three = WorksheetFunction.Match(gr_vol, ws.Cells("L2:L" & lastRow), 0)
ticker_three = ticker_three + 1
gr_total = ws.Cells(ticker_three, 9).Value

'Print the values in the correct cells
ws.Cells(2, 17).Value = gr_per_inc
ws.Cells(3, 17).Value = gr_per_dec
ws.Cells(4, 17).Value = gr_vol

Next ws


End Sub

