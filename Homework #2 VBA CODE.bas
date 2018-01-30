Attribute VB_Name = "Module1"
Sub market()

Dim ws As Worksheet
For Each ws In Worksheets

ws.Cells(1, 8).Value = "Ticker"
ws.Cells(1, 9).Value = "Total Stock Volume"

Dim ticker_name As String
Dim ticker_total As Double
Dim summary_table_row As Integer
Dim lastrow As Long

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row




 summary_table_row = 1
 
 
 
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker_name = ws.Cells(i, 1).Value
    
        ticker_total = ticker_total + ws.Cells(i, 7).Value
    
        summary_table_row = summary_table_row + 1
    
        ws.Range("h" & summary_table_row).Value = ticker_name
        ws.Range("I" & summary_table_row).Value = ticker_total

        Brand_total = 0

        Brand_total = Brand_total + ws.Cells(i, 7).Value
        
  
    End If
    Next i
 Next ws
 
 
 
End Sub
