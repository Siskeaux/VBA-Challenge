Attribute VB_Name = "Module2"
Sub MarketAnalysis():

    For Each ws In Worksheets
    
        Dim worksheet_name As String
        'Declaring data-types as either Long or Double, depending on their purposes, for proper viewing.
        Dim i As Long
        Dim j As Long
        Dim TickerCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PercentageChange As Double
        'Bonus Stuff, declarations.
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        'Giving worksheet_name a shortform.
        worksheet_name = ws.Name
        
        'Defining Column Headers as needed for the assignment.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        'Bonus stuff, column headers
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Setting Ticker-Counter to the first row.
        TickerCount = 2
        
        'Setting the beginning row as 2.
        j = 2
        
        'Find the last non-blank cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows
            For i = 2 To LastRowA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I (#9)
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate and write Yearly Change in column J (#10)
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TickerCount, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentageChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickerCount, 11).Value = Format(PercentageChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume in column L (#12)
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickerCount by 1
                TickerCount = TickerCount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Locating the final populated cell in the I column.
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Prep for Summary.
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            'For-Loop Summary.
            For i = 2 To LastRowI
            
                'Bonus stuff, to get greatest total volume, compare initial value with following value, if it is larger, then repopulate cells with the larger value, and check the next value, and so on.
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'Bonus stuff, to get greatest total increase, same logic as greatest total volume.
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                'Bonus stuff, for greatest decrease, same logic as previous two, but check to see if the initial value is smaller, rather than larger.
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            'Bonus stuff, properly formatting the summary results in ws.Cells.
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            
        'Automagically adjusting columns to fit without issues.
        Worksheets(worksheet_name).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub

