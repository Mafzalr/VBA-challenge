Attribute VB_Name = "Module1"
Sub yearly_stock_analysis()

'declaring the variables
    Dim ticker, prev_ticker As String
    Dim yearly_change, yearly_change_last, percentage_change, percentage_change_last, counter As Long
    Dim first_open_val, first_close_val, last_close_val, first_high_val, first_low_val, i_val As Long
    Dim total_stock_volume As LongLong
    Dim increase_ticker, decrease_ticker, volume_ticker As String
    Dim increase_value, decrease_value, volume_value As Long
    Dim min_value, max_value, max_volume_value As Double
         
    For Each ws In Worksheets
'initializing the variable
                min_value = 0
                max_value = 0
                max_volume_value = 0
                total_stock_volume = 0
'last row of the sheet
                last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
'assigning the headers
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
                ws.Cells(2, 16).Value = "Greatest % Increase"
                ws.Cells(3, 16).Value = "Greatest % Decrease"
                ws.Cells(4, 16).Value = "Greatest Total Volume"
                ws.Cells(1, 17).Value = "Ticker"
                ws.Cells(1, 18).Value = "Value"
'initializing the counter value
                counter = 2
'initializing the first_open_value
                first_open_val = ws.Cells(2, 3)
                
                'MsgBox (last_row)
                
'first for loop to get the counter value and the distinct ticker values
                For i = 2 To last_row

'                    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
'                        counter = counter + 1
'                        ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
'                        first_open_val = ws.Cells(i, 3).Value
'                        i_val = i
'                    End If

                    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
                        total_stock_volume = total_stock_volume + ws.Cells(i, 7)
                    Else
                        last_close_val = ws.Cells(i, 6).Value
                        yearly_change = last_close_val - first_open_val
                        ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
                        ws.Cells(counter, 10).Value = yearly_change
                        ws.Cells(counter, 11).Value = Format((yearly_change / first_open_val), "Percent")
                        ws.Cells(counter, 12).Value = total_stock_volume
                        
                        total_stock_volume = 0
                        counter = counter + 1
                        first_open_val = ws.Cells(i + 1, 3).Value
                        
                    End If

                    If ws.Cells(counter, 10).Value > 0 Then
                        ws.Cells(counter, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(counter, 10).Interior.ColorIndex = 3
                    End If
                    
                Next i
                
                max_volume = ws.Cells(2, 11).Value
                min_volume = ws.Cells(2, 11).Value
                max_volume_value = ws.Cells(2, 12).Value
                
'second for loop where the counter value is used, here counter will work as my last row  distinct ticker values


                For j = 3 To counter
                
                    If ws.Cells(j, 11) > max_value Then
                        max_value = ws.Cells(j, 11).Value
                        increase_ticker = ws.Cells(j, 9).Value
                    End If
                    
                    If ws.Cells(j, 11) < min_value Then
                        min_value = ws.Cells(j, 11).Value
                        decrease_ticker = ws.Cells(j, 9).Value
                    End If
                    
                    If ws.Cells(j, 12) > max_volume_value Then
                        max_volume_value = ws.Cells(j, 12).Value
                        volume_ticker = ws.Cells(j, 9).Value
                    End If
                   
                
                Next j
                 
                'display the greatest results
                ws.Cells(2, 17).Value = increase_ticker
                ws.Cells(2, 18).Value = Format(max_value, "Percent")
                ws.Cells(3, 17).Value = decrease_ticker
                ws.Cells(3, 18).Value = Format(min_value, "Percent")
                ws.Cells(4, 17).Value = volume_ticker
                ws.Cells(4, 18).Value = max_volume_value
                

    Next ws

End Sub



 

