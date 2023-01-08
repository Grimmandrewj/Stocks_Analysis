Sub Stocks()

'Instructions - Create script that loops through stocks and outputs the following:
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
    'Use conditional formatting to highlight positive change in green and negative in red
    
'Additional/Bonus
    'Calculate and print Greatest increase, decrease, and volume in each worksheet
    'Ensure script works on all worksheets
        

    'Apply script for each Worksheet
    For Each ws In Worksheets

        'Variable for worksheet name
        Dim WorksheetName As String
        
        'Currently referenced row
        Dim i As Double
        
        'Starting row of summary
        Dim j As Double
        
        'Variable for summary row
        Dim Summary_Row As Double
        
        'Variables for summary values
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Volume As Double
        
        'Variables for additional summary values
        Dim Greatest_Inc As Double
        Dim Greatest_Dec As Double
        Dim Greatest_Vol As Double
        
        'Grab worksheet name
        WorksheetName = ws.Name
        
        'Create column header values in each worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set summary row to first row beneath headers
        Summary_Row = 2
        
        'Set starting row to 2
        j = 2
        
        'Set initial total ticker volume
        Total_Volume = 0
        
        'Find the last row with a value in column A
        Last_A = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows in A
            For i = 2 To Last_A
            
                'Check if next row is the same.  If not:
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    'Get Ticker Name
                    TickerName = ws.Cells(i, 1).Value
                
                    'Print Ticker Name to summary column
                    ws.Cells(Summary_Row, 9).Value = TickerName
                    
                    'Calculate and print yearly change to summary column
                    ws.Cells(Summary_Row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    
                        'Set conditional formatting
                        If ws.Cells(Summary_Row, 10).Value < 0 Then
                        
                            'Set fill color to red if negative
                            ws.Cells(Summary_Row, 10).Interior.ColorIndex = 3
                        
                        ElseIf ws.Cells(Summary_Row, 10).Value > 0 Then
                        
                            'Set fill color to green if positive
                            ws.Cells(Summary_Row, 10).Interior.ColorIndex = 4
                            
                        Else
                        
                            'Set fill color to gray if neither positive nor negative
                            ws.Cells(Summary_Row, 10).Interior.ColorIndex = 15
                            
                        End If
                        
                    'Calculate and print percent change to summary column
                    Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(Summary_Row, 11).Value = Percent_Change
                                      
                    'Calculate and print total volume to summary column
                    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                    ws.Cells(Summary_Row, 12).Value = Total_Volume
                                
                    'Move to next row in Summary column
                    Summary_Row = Summary_Row + 1
                    
                    'Raise starting row in summary by 1
                    j = i + 1
                  
                    ' Reset the Total Volume
                    Total_Volume = 0

                ' If the cell below is the same ticker:
                Else

                    ' Add to the Total Volume
                    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                        
                
                End If
                                
            Next i
            
        'Find last row with a value in column I
        Last_I = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Set values for additional summary columns
        Greatest_Inc = ws.Cells(2, 11).Value
        Greatest_Dec = ws.Cells(2, 11).Value
        Greatest_Vol = ws.Cells(2, 12).Value
        
            'Loop through all rows in I
            For i = 2 To Last_I
            
                'Print increase and overwrite if next increase is higher
                If ws.Cells(i, 11).Value > Greatest_Inc Then
                    Greatest_Inc = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
                Else
                
                    Greatest_Inc = Greatest_Inc
                    
                End If
                    
                'Print decrease and overwrite if next increase is lower
                If ws.Cells(i, 11).Value < Greatest_Dec Then
                    Greatest_Dec = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    
                Else
                
                    Greatest_Dec = Greatest_Dec
                    
                End If
                
                'Print volume and overwrite if next volume is higher
                If ws.Cells(i, 12).Value > Greatest_Vol Then
                    Greatest_Vol = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    
                Else
                
                    Greatest_Vol = Greatest_Vol
                    
                End If
                
            'Print values of additional summaries in rows
            ws.Cells(2, 17).Value = Greatest_Inc
            ws.Cells(3, 17).Value = Greatest_Dec
            ws.Cells(4, 17).Value = Greatest_Vol
                
            Next i
       
    Next ws


End Sub
