Sub Ticker()

'Define everything
Dim Ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'loop over each worksheet in the workbook
For Each ws In Worksheets
    
    'Make the worksheet active
    ws.Activate

'Find the last row in each worksheet
lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row


'Crease Header Csolumns for each worksheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Value"
    
    
'Variables for each ws
number_tickers = 0
Ticker = ""
yearly_change = 0
year_open = 0
total_stock_volume = 0


'create loop for rows
        ' Loop through all ticker names skipping header row
              For i = 2 To lastRowState
              
                 ''Value of the ticker
                 Ticker = Cells(i, 1).Value
                 
                 'Calc year open price
                 If year_open = 0 Then
                    year_open = Cells(i, 3).Value
                End If
                
                 'Add total stock volume values for tickers
                 total_stock_volume = total_stock_volume + Cells(i, 7).Value
                 
                 'Tell it to run this if we get a different ticker in the lsit
                 If Cells(i + 1, 1).Value <> Ticker Then
                    number_tickers = number_tickers + 1
                    Cells(number_tickers + 1, 9) = Ticker
                
                        
                        'get closing price
                        year_close = Cells(i, 6).Value
                        
                        
                        'calculate yearly change and percent change
                        yearly_change = year_close - year_open
                       
                        
                        ' Add yearly change value to the appropriate cell in each worksheet.
                         Cells(number_tickers + 1, 10).Value = yearly_change
                         
                        ' If yearly change value is greater than 0, shade cell green.
                        If yearly_change > 0 Then
                            Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
                            
                       ' If yearly change value is less than 0, shade cell red.
                         ElseIf yearly_change < 0 Then
                            Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
                        
                         ' If yearly change value is 0, shade cell yellow.
                        Else
                            Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
                        End If
                        
                        ' Calculate percent change value for ticker.
                        If year_open = 0 Then
                             percent_change = 0
                        Else
                             percent_change = (yearly_change / year_open)
                        End If
                        ' Format the percent_change value as a percent.
                        Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
                        
                        
                        ' Format the percent_change value as a percent.
                            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
                        
                        ' Set opening price back to 0 when we get to a different ticker in the list.
                            opening_price = 0
            
                        'Add total stock volume value to the appropriate cell in each worksheet.
                            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
                        'Set total stock volume back to 0 when we get to a different ticker in the list.
                            total_stock_volume = 0
                End If
                
            Next i
            
           Next ws

        
End Sub
