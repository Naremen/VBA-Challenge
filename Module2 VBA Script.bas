Attribute VB_Name = "Module1"

Sub Tryout()

Dim ws As Worksheet


For Each ws In Worksheets

    Const TICKER_COL As Integer = 1
    'Set variables
    Dim ticker_name As String
    Dim volume As Long
    
    Dim open_price As Double
    Dim close_price As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim summary_table_row As Long
    Dim LastRow As Long
    volume = 0
    summary_row = 2
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total_volume = 0
    x = 0
    YearlyChange = 0
    
    
    
    
    'Column Headings
    ws.Range("L1").Value = "Ticker"
    ws.Range("M1").Value = "YearlyChange"
    ws.Range("N1").Value = "PercentChange"
    ws.Range("O1").Value = "volume"
    
    'Define last row in the worksheet
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
    
    'Loop Through All Rows
    For i = 2 To LastRow
    
        ticker_name = ws.Cells(i, TICKER_COL).Value
        volume = volume + ws.Cells(i, 7).Row
               
        'Last row of current stock
        If ws.Cells(i + 1, TICKER_COL).Value <> ticker_name Then
           
            'Inputs
            close_price = ws.Cells(i, 6).Value
            open_price = ws.Cells(summary_row, 3).Value
                     
            'Calculations
            YearlyChange = close_price - open_price
        If open_price <> 0 Then
            PercentChange = (YearlyChange / open_price)
        Else
            PercentageChange = 0
        
        End If
                        
            'Outputs
            ws.Range("L" & summary_row).Value = ticker_name
            ws.Range("O" & summary_row).Value = volume
            ws.Range("M" & summary_row).Value = YearlyChange
            ws.Range("N" & summary_row).Value = PercentChange
            ws.Range("N" & summary_row).Style = "Percent"
            ws.Range("N" & summary_row).NumberFormat = "0.00%"
                    
            'Prepare for next stock
            'open_price = ws.Cells(summary_row, 3).Value
            YearlyChange = 0
            volume = 0
            summary_row = summary_row + 1
        
        
        End If
    
    Next i
    
        For i = 2 To LastRow
    
        If ws.Range("M" & i).Value > 0 Then
            ws.Range("M" & i).Interior.ColorIndex = 4
        
        ElseIf ws.Range("M" & i).Value < 0 Then
            ws.Range("M" & i).Interior.ColorIndex = 3
        
        End If
    
    
        If ws.Range("N" & i).Value > 0 Then
            ws.Range("N" & i).Interior.ColorIndex = 4
        
        ElseIf ws.Range("N" & i).Value < 0 Then
            ws.Range("N" & i).Interior.ColorIndex = 3
        
        End If
    
Next i

    Dim Greatestincrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    
    ws.Range("T1").Value = "Ticker"
    ws.Range("U1").Value = "Value"
    ws.Range("S2").Value = "GreatestPercentIncrease"
    ws.Range("S3").Value = "GreatestPercentDecrease"
    ws.Range("S4").Value = "GreatestTotalVolume"
    
    For i = 2 To LastRow
        If ws.Range("N" & i).Value > ws.Range("U2").Value Then
                    ws.Range("U2").Value = ws.Range("N" & i).Value
                    ws.Range("T2").Value = ws.Range("L" & i).Value
                End If
        If ws.Range("N" & i).Value < ws.Range("U3").Value Then
                    ws.Range("U3").Value = ws.Range("N" & i).Value
                    ws.Range("T3").Value = ws.Range("L" & i).Value
                End If
        If ws.Range("O" & i).Value > ws.Range("U4").Value Then
                    ws.Range("U4").Value = ws.Range("O" & i).Value
                    ws.Range("T4").Value = ws.Range("L" & i).Value
                End If
        
Next i

Next ws
    
End Sub
