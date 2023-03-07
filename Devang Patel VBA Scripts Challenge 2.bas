Sub StockAnalysis()

    Dim unique_ticker_count As Integer
    Dim row_num As Long
    Dim ticker_name As String
    Dim previous_ticker_name As String
    Dim stock_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Variant
    Dim max_incr(1 To 2) As Variant
    Dim max_decr(1 To 2) As Variant
    Dim max_vol(1 To 2) As Variant
    Dim start_time As Double
    Dim seconds_elapsed As Double
    
    For Each ws In Worksheets

        With ws.Sort
             .SortFields.Add Key:=Range("A1"), Order:=xlAscending
             .SortFields.Add Key:=Range("B1"), Order:=xlAscending
             .SetRange Columns("A:G")
             .Header = xlYes
             .Apply
        End With
        
        ws.Range("I1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"

        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        ReDim stock_volume_array(9999, 3) As Variant
        
        unique_ticker_count = 0
        row_num = 2
        ticker_name = ws.Range("A" & row_num)
        stock_volume_array(0, 0) = ticker_name
        stock_volume = CDbl(ws.Range("G" & row_num).Value)
        open_price = CDbl(ws.Range("C" & row_num).Value)
        
        previous_ticker_name = ticker_name
        
        max_incr(1) = ""
        max_decr(1) = ""
        max_vol(1) = ""
        max_incr(2) = 0
        max_decr(2) = 0
        max_vol(2) = 0
        
        Do While (ticker_name <> "")
            
            If ticker_name = previous_ticker_name Then
                
                stock_volume_array(unique_ticker_count, 1) = stock_volume_array(unique_ticker_count, 1) + stock_volume
            
            Else

                close_price = ws.Range("F" & row_num - 1)
                
                yearly_change = close_price - open_price
                
                stock_volume_array(unique_ticker_count, 2) = yearly_change
                
                If open_price <> 0 Then
                    percent_change = (close_price - open_price) / open_price
                Else
                    percent_change = "n/a"
                End If
                
                 stock_volume_array(unique_ticker_count, 3) = percent_change
                 
                If percent_change > max_incr(2) And percent_change <> "n/a" Then
                    max_incr(1) = ticker_name
                    max_incr(2) = percent_change
                End If
                
                If percent_change < max_decr(2) Then
                    max_decr(1) = ticker_name
                    max_decr(2) = percent_change
                End If
                
                unique_ticker_count = unique_ticker_count + 1
                
                stock_volume_array(unique_ticker_count, 0) = ticker_name
                stock_volume_array(unique_ticker_count, 1) = stock_volume
                
                open_price = ws.Range("C" & row_num)
                
            End If
            
            previous_ticker_name = ticker_name
            
            row_num = row_num + 1
            ticker_name = ws.Range("A" & row_num)
            stock_volume = CDbl(ws.Range("G" & row_num).Value)

        Loop
        
        For i = 0 To unique_ticker_count - 1
            
            ticker_name = stock_volume_array(i, 0)
            ws.Range("I" & i + 2) = ticker_name
            
            stock_volume = stock_volume_array(i, 1)
            ws.Range("L" & i + 2) = stock_volume
            If stock_volume > max_vol(2) Then
                max_vol(1) = ticker_name
                max_vol(2) = stock_volume
            End If
            
            yearly_change = stock_volume_array(i, 2)
            ws.Range("J" & i + 2) = yearly_change
            
            If yearly_change > 0 Then
                ws.Range("J" & i + 2).Interior.Color = vbGreen
            Else
                ws.Range("J" & i + 2).Interior.Color = vbRed
            End If
            
            ws.Range("K" & i + 2) = stock_volume_array(i, 3)
        Next i

        ws.Range("P2") = max_incr(1)
        ws.Range("Q2") = max_incr(2)
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P3") = max_decr(1)
        ws.Range("Q3") = max_decr(2)
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P4") = max_vol(1)
        ws.Range("Q4") = max_vol(2)
        ws.Range("K2:K" & (2 + unique_ticker_count)).NumberFormat = "0.00%"
        
    Next ws
    
End Sub