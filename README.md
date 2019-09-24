# Homework
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
