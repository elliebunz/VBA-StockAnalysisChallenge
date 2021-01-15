Sub stockanalysis()
     
    'Establish all variables
    
    Dim ws As Worksheet
    Dim Percent_Change As Double
    Dim Volume_Total As Double
    Dim Ticker_Name As String
    Dim Annual_Change As Single
    Dim StartValue As Long
    Dim LastRow As Long
    Dim i As Long
    Dim j As Long

    
    'make sure to loop through all worksheets
    
    For Each ws In Worksheets
    
    'Set values - this is for every worksheet
    
        j = 0
        Volume_Total = 0
        Yearly_Change = 0
        StartValue = 2
        
     'create headers
        
        ws.Cells(1, 9).Value = "Stock Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'identify last row
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow

                'Check if within the same stock ticker
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Add to the Volume Total
                    
                    Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                    
                    'Account for zero total volume
                    
                    If Volume_Total = 0 Then
                    
                        'print the results
                        ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = 0
                        ws.Range("K" & 2 + j).Value = "%" & 0
                        ws.Range("L" & 2 + j).Value = 0
    
                'If the next cell has the same ticker,
                
                Else
                
                    If ws.Cells(StartValue, 3) = 0 Then
                    
                        For find_value = StartValue To i
                        
                            If ws.Cells(find_value, 3).Value <> 0 Then
                            
                                    StartValue = find_value
                                    
                                    Exit For
                                    
                            End If
                            
                        Next find_value
                        
                    End If
                    
                    'Calculate Annual and Percent Changes
                    
                    Annual_Change = (ws.Cells(i, 6) - ws.Cells(StartValue, 3))
                    Percent_Change = Round((Annual_Change / ws.Cells(StartValue, 3) * 100), 2)
                    
                    'Start new ticker
                    
                    StartValue = i + 1
                    
                    'print results to a new worksheet
                    
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(Annual_Change, 2)
                    ws.Range("K" & 2 + j).Value = "%" & Percent_Change
                    ws.Range("L" & 2 + j).Value = Volume_Total
                    
                    'change colors of cells to show increase green and decrease red
                    
                    Select Case Yearly_Change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If
                   
                'reset variables for stock ticker
                
                Annual_Change = 0
                Volume_Total = 0
                j = j + 1
                
            'If ticker is still the same calculate sum of results
            
            Else
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
           End If
                            
        Next i
    Next ws
    
End Sub

