Attribute VB_Name = "Module1"
'The VBA code below is designed to do the following:
'Create a script that will loop through all the stocks for one year and output the following information:
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.


Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'We need to find the current row
        Dim i As Long
        'Kick the process
        Dim j As Long
        'Counter to adjust ticker
        Dim TickCount As Long
        'We need to get the lastrow column A
        Dim LastRowA As Long
        'We need tp get the last row column I
        Dim LastRowI As Long
        'Creating a variable for the prcent change
        Dim PerChange As Double
        'Creating a variable for the  increase calculation
        Dim GreatIncr As Double
        'Creating a variable for the  greatest decrease calculation
        Dim GreatDecr As Double
        'Creating a variable for the greatest total volume
        Dim GreatVol As Double
        'We have now set all the variable that we need
        
        'Get the WorksheetName (In this case we have 3 worksheet)
        WorksheetName = ws.Name
        
        'Create column headers as per the instruction
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'We need to set the counter
        TickCount = 2
        
        'Set start row to 2 since row 1 is the header
        j = 2
        
        'We need to determine the last row of data and find the blank
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
            'Loop through all rows as per the instruction
            For i = 2 To LastRowA
            
                'Verify that the ticket name change
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I (#9) as per the isntruction
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
                'We need to calculate and write Yearly Change in column J as per the instruction (#10)
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Adding an if statement
                    If ws.Cells(TickCount, 10).Value < 0 Then
                
                    'Changing the background color of the cell to red
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Changing the background color of the cell to green
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Formatting the percentage
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'We need to calculate and write total volume in column L (#12)
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1 to move to the next
                TickCount = TickCount + 1
                
                'Set new start row of the ticker block so that we dont have issue
                'this is actually a good approach
                j = i + 1
                
                End If
            
            Next i
            
        'Lookup the last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
        'Summarizing the data
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            'Loop for the summary of the data
            For i = 2 To LastRowI
            
              
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            'Writing the summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            
       
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
