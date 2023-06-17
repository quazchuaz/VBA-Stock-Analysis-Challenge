Sub StockAnalysis()
Dim i As Long, lastrow As Long

Dim stock_start As Double, stock_end As Double, total_vol As Double, perchange As Double, perchange_high As Double, perchange_low As Double, total_vol_high As Double

Dim ticker As String, perchange_high_ticker As String, perchange_low_ticker As String, total_vol_high_ticker As String

Dim summary_row As Integer

Dim ws As Worksheet
    
For Each ws In Worksheets

lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest Percent Increase"
ws.Range("O3").Value = "Greatest Percent Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
        
summary_row = 2
ticker = ws.Range("A2")
stock_start = ws.Range("C2")
total_vol = 0
perchange_high = 0
perchange_low = 0
total_vol_high = 0
perchange_high_ticker = ""
perchange_low_ticker = ""
total_vol_high_ticker = ""
        
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(summary_row, 9).Value = ticker
stock_end = ws.Cells(i, 6).Value
ws.Cells(summary_row, 10).Value = (stock_end - stock_start)
perchange = ((stock_end - stock_start) / stock_start)
total_vol = total_vol + ws.Cells(i, 7).Value
ws.Range("I" & summary_row).Value = ticker
ws.Range("J" & summary_row).Value = (stock_end - stock_start)
ws.Range("K" & summary_row).Value = perchange
ws.Range("L" & summary_row).Value = total_vol
ws.Range("K" & summary_row).NumberFormat = "0.00%"
summary_row = summary_row + 1
stock_end = 0
total_vol = 0
stock_start = ws.Cells(i + 1, 3).Value
ticker = ws.Cells(i + 1, 1).Value
            
Else
total_vol = total_vol + ws.Cells(i, 7).Value

End If

Next i
        
Dim cell As Range
For Each cell In ws.Range("J2:J" & lastrow)

If cell.Value > 0 Then
cell.Interior.ColorIndex = 4

ElseIf cell.Value < 0 Then

cell.Interior.ColorIndex = 3
End If

Next cell
        
perchange_high = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
perchange_low = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
total_vol_high = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
        
ws.Range("Q2").Value = perchange_high
ws.Range("Q3").Value = perchange_low
ws.Range("Q4").Value = total_vol_high
        
For i = 2 To lastrow
If ws.Cells(i, 11).Value = perchange_high Then
perchange_high_ticker = ws.Cells(i, 9).Value
ws.Range("P2").Value = perchange_high_ticker

ElseIf ws.Cells(i, 11).Value = perchange_low Then
perchange_low_ticker = ws.Cells(i, 9).Value
ws.Range("P3").Value = perchange_low_ticker

ElseIf ws.Cells(i, 12).Value = total_vol_high Then
total_vol_high_ticker = ws.Cells(i, 9).Value
ws.Range("P4").Value = total_vol_high_ticker

End If

Next i
        
ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

End Sub