# Homework
#Assignment Easy, Moderate, and Hard

Sub stockdata():
    'assignment easy
    'multiple_year_stock_data
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim summary_table_row As Integer
    summary_table_row = 2

    For Each ws In Worksheets
        ws.Activate
        summary_table_row = 2
        ws.Range("i1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            For i = 2 To lastrow
                

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'set the ticker name
                Ticker = Cells(i, 1).Value
                
                'add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

                'print the ticker name in the summary table
                ws.Range("i" & summary_table_row).Value = Ticker
                
                'print the total stock volume to the summary table
                ws.Range("j" & summary_table_row).Value = Total_Stock_Volume

                'add one to the summary table row
                summary_table_row = summary_table_row + 1

                'reset the total stock volume
                Total_Stock_Volume = 0

            Else

                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

            End If
                
            Next i
    Next ws
        
End Sub

#Assignment Moderate & Hard
Sub stockdata():

'assignment moderate
'multiple_year_stock_data

Dim ws As Worksheet
Dim Ticker As String
Dim yearly_change As Double
'yearly_change = 0
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
Dim summary_table_row As Double
summary_table_row = 2
Dim year_open As Double
Dim year_close As Double
Dim percent_change As Double

'percent_change = 0

    For Each ws In Worksheets
    ws.Activate
    
    summary_table_row = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set year open price
    year_open = Range("C2").Value
    
        For i = 2 To lastrow
            'If ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set the ticker name
            Ticker = Cells(i, 1).Value
    
            'add to the total stock volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            'year_open = Cells(i, 3).Value
            year_close = Cells(i, 6).Value
                
            'add to the yearly Change
            yearly_change = year_close - year_open
            ws.Cells(i, 11).NumberFormat = "0.00%"
            percent_change = ((year_close - year_open) / year_open) * 100
            
            'year_close = ws.Range("F" & summary_table_row).Value
                                   
            'If yearly_open = 0 And yearly_close = 0 Then
                'percent_change = 0
                
           ' Else: yearly_open = 0 And yearly_close <> 0
                'percent_change = 1
                
                 
            'End If
            
            'print the ticker name in the summary table
            ws.Range("I" & summary_table_row).Value = Ticker
            
            
            'print the total stock volume to the summary table
            ws.Range("L" & summary_table_row).Value = Total_Stock_Volume
            
            'print the yearly change
            ws.Range("J" & summary_table_row).Value = yearly_change
            
            
            'print the percent change
            ws.Range("K" & summary_table_row).Value = percent_change
            
            'add one to the summary table row
            summary_table_row = summary_table_row + 1
            
            'reset the total stock volume
            Total_Stock_Volume = 0
            
            'reset open price
            year_open = ws.Cells(i + 1, 3)
            
            Else
    
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            End If
        
        Next i
            
                ylastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
                
                'color cell green for positive and red for negative yearly change
                For j = 2 To ylastrow
                
                If ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0 Then
                
                ws.Cells(j, 10).Interior.Color = vbGreen
                
                Else
                
                ws.Cells(j, 10).Interior.Color = vbRed
                
                
                End If
                
                
                Next j
            
            
            'Greatest % increase, Greatest % decrease, and Greatest total volume
            
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            
            For k = 2 To ylastrow
            
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & ylastrow)) Then
            
            ' ws.range("P2").value = ws.Range("I" & summary_table_row).Value
            ' ws.range("Q2").value = ws.Range("K" & summary_table_row).Value
            
            ws.Range("P2").Value = ws.Cells(k, 9).Value
            ws.Range("Q2").Value = ws.Cells(k, 11).Value
            ws.Range("Q2").NumberFormat = "0.00%"
            
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & ylastrow)) Then
            
            ' ws.range("P3").value = ws.Range("I" & summary_table_row).Value
            ' ws.range("Q3").value = ws.Range("K" & summary_table_row).Value
            
            ws.Range("P3").Value = ws.Cells(k, 9).Value
            ws.Range("Q3").Value = ws.Cells(k, 11).Value
            ws.Range("Q3").NumberFormat = "0.00%"
            
            ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & ylastrow)) Then
            
            ' ws.range("P4").value = ws.Range("I" & summary_table_row).Value
            ' ws.range("Q4").value = ws.Range("L" & summary_table_row).Value
            
            ws.Range("P4").Value = ws.Cells(k, 9).Value
            ws.Range("Q4").Value = ws.Cells(k, 12).Value
            
            'add one to the summary table row
            'summary_table_row = summary_table_row + 1
            
            End If
            
            
            Next k
    
            
    Next ws
                
            
End Sub
