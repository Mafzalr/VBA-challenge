Attribute VB_Name = "Module1"
Sub yearly_stock_analysis()

 For Each ws In Worksheets
                
'declaring the variables

                Dim ticker, prev_ticker As String
                Dim yearly_change, yearly_change_last, percentage_change, percentage_change_last, counter As Long
                Dim first_open_val, first_close_val, last_close_val, first_high_val, first_low_val, i_val As Long
                Dim total_stock_volume As Long
                Dim increase_ticker, decrease_ticker, volume_ticker As String
                Dim increase_value, decrease_value, volume_value As Long
                Dim min_value, max_value, max_volume_value As Double
                '<ticker>    <date>  <open>  <high>  <low>   <close> <vol>
                
                min_value = 0
                max_value = 0
                max_volume_value = 0
                
'last row of the sheet
                
                last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
                total_stock_volume = 0
                
                
                'MsgBox last_row
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
                
'first for loop to get the counter value and the distinct ticker values
                
                counter = 1
                
                For i = 2 To last_row
                    
                
                    'If Cells(i, 1).Value <> "ABKV" Then
                
                    'The total stock volume of the stock
                    
                    
                    'total_stock_volume = CLng(total_stock_volume) + Cells(i, 7).Value
                    
                    
                    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                    
                        counter = counter + 1
                        
                        ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
                        'last_close_val = Cells(i - 1, 1).Value
                        
                        first_open_val = ws.Cells(i, 3).Value
                        
                        'Cells(counter, 12).Value = total_stock_volume
                        
                        
                    
                        i_val = i
                    
                    'Else
                        
                        'total_stock_volume = clng(total_stock_volume) + Cells(i, 7).Value
                        
                        
                    
                    End If
                        
                     If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                    
                        
                       ' https://www.techonthenet.com/excel/formulas/format_string.php
                      ' https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat
                        
                        
                        last_close_val = ws.Cells(i, 6).Value
                        
                        
                        
                    'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
                
                       ws.Cells(counter, 10).Value = last_close_val - first_open_val
                       
                     'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year
                    
                    ws.Cells(counter, 11).Value = Format((ws.Cells(counter, 10).Value / first_open_val), "Percent")
                    
                      ws.Cells(counter, 12).Value = CLng(total_stock_volume)
                        
                        'total_stock_volume = Cells(i, 7).Value
                        
                    
                    
                    End If
                    
                    
                     If ws.Cells(counter, 10).Value > 0 Then
                       
                        ws.Cells(counter, 10).Interior.ColorIndex = 4
                       
                       Else
                       
                        ws.Cells(counter, 10).Interior.ColorIndex = 3
                       
                     End If
                    
                    
                    'End If
                    
                    
                Next i
                    
                    
 
'nested for loop where the counter value is to get the counter value and the distinct ticker values
                    
                For j = 2 To counter
                
                
                    For k = 2 To last_row
                    
                        If ws.Cells(j, 9).Value = ws.Cells(k, 1).Value Then
                        
                            ws.Cells(j, 12).Value = ws.Cells(j, 12).Value + ws.Cells(k, 7).Value
                            
                        
                        End If
                        
                    
                    Next k
                    
                    
                    If ws.Cells(j, 11).Value > max_value Then
                        
                        max_value = ws.Cells(j, 11).Value
                        
                        ws.Cells(2, 18).Value = Format(max_value, "Percent")
                        ws.Cells(2, 17).Value = ws.Cells(j, 9).Value
                        
                        
                    ElseIf ws.Cells(j, 11).Value < min_value Then
                        
                        min_value = ws.Cells(j, 11).Value
                        ws.Cells(3, 18).Value = Format(min_value, "Percent")
                        ws.Cells(3, 17).Value = ws.Cells(j, 9).Value
                        
                    End If
                    
                    If ws.Cells(j, 12).Value > max_volume_value Then
                    
                        max_volume_value = ws.Cells(j, 12).Value
                    
                        ws.Cells(4, 18).Value = max_volume_value
                        ws.Cells(4, 17).Value = ws.Cells(j, 9).Value
                    
                    End If
                
                
                Next j
                
                
                    
Next ws

End Sub



