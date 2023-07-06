Attribute VB_Name = "Module1"
Sub Button1_Click()
     Dim ws As Worksheet
    Dim ticker As String
    Dim i As Long
    Dim summary_table_row As Long
    Dim lastRow As Long
    Dim total_stock_volume As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim start As Long
    
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String

    summary_table_row = 2
    max_increase = 0
    max_decrease = 0
    max_volume = 0
    
    For Each ws In Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        total_stock_volume = 0
        start = 2
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                yearly_change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
                
                If ws.Cells(start, 3).Value <> 0 Then
                    percent_change = yearly_change / ws.Cells(start, 3).Value
                Else
                    percent_change = 0
                End If
                
                ws.Range("K" & summary_table_row).Value = percent_change
                ws.Range("J" & summary_table_row).Value = yearly_change
                ws.Range("L" & summary_table_row).Value = total_stock_volume
                ws.Range("I" & summary_table_row).Value = ticker
                
                If yearly_change >= 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
                
                If percent_change > max_increase Then
                    max_increase = percent_change
                    max_increase_ticker = ticker
                End If

                If percent_change < max_decrease Then
                    max_decrease = percent_change
                    max_decrease_ticker = ticker
                End If

                If total_stock_volume > max_volume Then
                    max_volume = total_stock_volume
                    max_volume_ticker = ticker
                End If
                
                summary_table_row = summary_table_row + 1
                total_stock_volume = 0
                yearly_change = 0
                ticker = ""
                start = i + 1
            Else
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ws.Range("Q2").Value = max_increase_ticker
        ws.Range("Q3").Value = max_decrease_ticker
        ws.Range("Q4").Value = max_volume_ticker
    
        ws.Range("R2").Value = max_increase
        ws.Range("R3").Value = max_decrease
        ws.Range("R4").Value = max_volume
        
        summary_table_row = 2
        max_increase = 0
        max_decrease = 0
        max_volume = 0
    Next ws
End Sub
