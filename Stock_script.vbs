Sub StockAnalysis(ws As Worksheet)
Dim total_volume As Double
Dim stock_name As String
Dim yearly_change As Double
Dim percent_change As Double
Dim opening As Double
Dim closing As Double
Dim last_row As Long
Dim total_row As Long

last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
stock_name = ws.Cells(2, 1)
opening = ws.Cells(2, 3)
total_volume = 0
total_row = 2

For i = 2 To last_row
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        total_volume = total_volume + ws.Cells(i, 7)
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        total_volume = total_volume + ws.Cells(i, 7)
        closing = ws.Cells(i, 6)
        yearly_change = closing - opening
        percent_change = (closing - opening) / opening
        ws.Cells(total_row, 9).Value = stock_name
        ws.Cells(total_row, 10).Value = yearly_change
        ws.Cells(total_row, 11).Value = percent_change
        ws.Cells(total_row, 12).Value = total_volume
        total_volume = 0
        opening = ws.Cells(i + 1, 3).Value
        stock_name = ws.Cells(i + 1, 1).Value
        total_row = total_row + 1
    End If
Next i
        
End Sub

Sub GreatestCheck(ws As Worksheet)
Dim last_row As Long
Dim stock_name As String
Dim stock_name_inc As String
Dim stock_name_dec As String
Dim total_volume As Double
Dim max_dec As Double
Dim max_inc As Double

last_row = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
total_volume = ws.Cells(2, 12).Value
max_inc = ws.Cells(2, 11).Value
max_dec = ws.Cells(2, 11).Value
stock_name = ws.Cells(2, 9).Value
stock_name_inc = ws.Cells(2, 9).Value
stock_name_dec = ws.Cells(2, 9).Value

For i = 2 To last_row
    If ws.Cells(i, 11).Value > max_inc Then
        max_inc = ws.Cells(i, 11).Value
        stock_name_inc = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 11).Value < max_dec Then
        max_dec = ws.Cells(i, 11).Value
        stock_name_dec = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 12).Value > total_volume Then
        total_volume = ws.Cells(i, 12).Value
        stock_name = ws.Cells(i, 9).Value
    End If
    Next i
    
    ws.Range("P2").Value = stock_name_inc
    ws.Range("P3").Value = stock_name_dec
    ws.Range("P4").Value = stock_name
    ws.Range("Q2").Value = max_inc
    ws.Range("Q3").Value = max_dec
    ws.Range("Q4").Value = total_volume
    
End Sub

Sub TableLabels(ws As Worksheet)
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Columns("I:Q").AutoFit
    ws.Columns("K").NumberFormat = "0.00%"
End Sub


Sub AllSheetsAnaylsis()
Dim ws As Worksheet
Dim cell As Range
For Each ws In ThisWorkbook.Sheets
    TableLabels ws
    StockAnalysis ws
    GreatestCheck ws
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "0"
    ws.Range("J:J").NumberFormat = "$#,##0.00"
    For Each cell In ws.Range("J:K")
        If cell.Value > 0 And cell.Value <> "Yearly Change" And cell.Value <> "Percent Change" Then
            cell.Interior.ColorIndex = 4
        ElseIf cell.Value < 0 Then
            cell.Interior.ColorIndex = 3
        End If
    Next cell
Next ws
End Sub


Sub ResetAll()
Dim ws As Worksheet
Dim clear As Range
For Each ws In ThisWorkbook.Sheets
    Set clear = ws.Range("I:Q")
    clear.ClearContents
    clear.ClearFormats
Next ws
End Sub
