Attribute VB_Name = "Module1"
' Author: Michelle Carvalho (UTOR-VIRT-DATA-PT-10-2022-U-LOLC)
' Sources: YouTube, Google, LinkedIn, SkillShare, Microsoft Excel 2019 VBA and Macros (Text)
'________________________________________________________________________________________________________________________________________
Sub Create_StockData_Report()
    
        For Each ws In ActiveWorkbook.Worksheets    'For Loop: Specify the range through which code will execute through each worksheet
                                                    'By defining "ws" as the range for the For Loop, "ws" needs to be used as a precursor to any other object in the code (ie. Cells, Columns)
                                                    'This will ensure that it is applied to every worksheet as the code loops through each one
                                                                              
          ' Declaring variables
        Dim SummaryTable As Variant
        Dim StartDateStockValue As Double
        Dim EndDateStockValue As Double
        Dim YearlyChange As Single  'as it will be a decimal
        Dim PercentageChange As Variant   'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/variant-data-type
        Dim VolumeTotal As Variant
        Dim i As Long			' Long will deal with the large size of the data file
        Dim lastrow As Long		' Long will deal with the large size of the data file
        
    
            ' Automate creating headings/labels on each spreadsheet for results and autofit cell width to contents
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Result"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Columns("K:P").EntireColumn.AutoFit      'formats the colums to autofit to the contents
        

            'Constants
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row    'find the last row of our data; supports differing row #'s in each spreadsheet (where a range won't work)

        SummaryTable = 2  'add results to summary table starting in Row 2 (excludes header)
        
        VolumeTotal = 0      'volume counter will always start off at zero and must be defined before the nested For Loop (kind of like a parameter by which the For Loop will operate)
      
'____________________________________________________________________________________________________________________________________________________________________________

        ' Nested For Loop to define the Range of your data set on each worksheet (from row 2 to lastrow)

            For i = 2 To lastrow    'defines the sequence of the loop from data starting in row 2
            
'____________________________________________________________________________________________________________________________________________________________________________
        
       ' Calculate the Percentage Change and Yearly Change (with formatting to show positive/negative change for Yearly Change)
        
        StartDateStockValue = ws.Cells(2, 3).Value  'specifies that the opening stock value will always be the first value (in column 3) for each new Ticker ID
        
        'Having now defined where to find the start date stock value, I need to define how to get the closing date stock value
              If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                
                EndDateStockValue = ws.Cells(i, 6).Value

                'Calculate the Yearly Change
                YearlyChange = EndDateStockValue - StartDateStockValue
                
                 'Calculate Percentage Change
                PercentageChange = (EndDateStockValue - StartDateStockValue) / StartDateStockValue

                ws.Cells(SummaryTable, 11).Value = YearlyChange    'adds the Yearly Change result to the summary table in column 11 (row 2)
                
                     'Add formatting
                If ws.Cells(SummaryTable, 11).Value < 0 Then       'conditional format for negative Yearly Change (red)
                    ws.Cells(SummaryTable, 11).Interior.ColorIndex = 3
                Else                                                  'conditional format for positive Yearly Change (green)
                    ws.Cells(SummaryTable, 11).Interior.ColorIndex = 4
                End If

                ws.Cells(SummaryTable, 12).Value = PercentageChange 'adds the Percentage Change result to the summary table in column 12 (row 2)
                
                ws.Cells(SummaryTable, 12).NumberFormat = "0.00%"    'formats cell as a percentage (https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage)
                
        
            End If
'____________________________________________________________________________________________________________________________________________________________________________

        'Calculate the Total Volume
              
              VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value  'sets the starting point for calculating a cumulative total for Volume
                                                                'In class example coded in the IF statement (as "Else"); when testing the results, the total excluded the last row from the tally
                                                                'After researching, I realized that in "Else", as it got to the end of a Ticker ID sequence, at the point where the ID changed
                                                                'it was holding/storing that value instead of including it (the "Else" was kind of working like a "stop here")...I think? (Still not 100% clear)
                                                                'After some trial & error, I realized an Else statement was not needed (though the code was needed), & worked when I coded to define the range of Volume
              
              If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then      'as it loops through each row, calculates a running total until ticker values are NOT the same (as in contiguous rows)
         
                        ws.Cells(SummaryTable, 10).Value = ws.Cells(i, 1).Value     'Display Ticker ID in Summary Table; it goes here, rather than the IF statement above because it is driven by the running count
    
                        ws.Cells(SummaryTable, 13).Value = VolumeTotal    'Display total volume in Summary Table

                        VolumeTotal = 0    'resets the volume total counter for the next Ticker ID
                        
                        SummaryTable = SummaryTable + 1     'pushes the next iteration to the next row in the Summary Table
            
              End If
                          
 '____________________________________________________________________________________________________________________________________________________________________________
            
            Next i  'move to next row in the For Loop
 '____________________________________________________________________________________________________________________________________________________________________________

            'BONUS: Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" (per worksheet)
            
            'Options that search for values <> zero won't work here, as it will simply return the first value that is <> zero and will stop (static operation)
            'Key: create a loop that builds on itself. First need to establish a starting value and begin by adding that baseline to the cells where the max/min will be once it completes all iterations.
            'By adding a value to the BONUS cells, as the For Loop moves through each iteration, it is constantly being replaced (making it a dynamic operation)
            'https://www.geeksforgeeks.org/how-to-use-for-next-loop-in-excel-vba/
            'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fornext-statement
                            'The ForÉNext statement syntax has these parts:
                            'counter: Required. Numeric variable used as a loop counter. (2 to last row)
                            'start: Required. Initial value of counter. (set the Bonus cells to 0)
                            'end: Required. Final value of counter. (Max, Min)
            
	    'Adds zero to the results table as a placeholder (in the cells where the max/min will be shown)	
            ws.Cells(2, 18).Value = 0
            ws.Cells(3, 18).Value = 0
            ws.Cells(4, 18).Value = 0
            
            'Looping through each row in the summary table, starting from zero, every value will be compared to each previous iteration.  In the first iteration, it will look for the first number higher
            'than zero and save/hold it.  In the next iteration, if the number is lower, it will skip to the next value, but if the next iteration is higher than the one currently saved, it will replace it.
            'It will repeat this process until it gets to the last used row.  Whatever number is saved at that point will then go into the summary table, along with the corresponding Ticker ID.
            
            For SummaryTable = 2 To lastrow
              If ws.Cells(SummaryTable, 12) > ws.Cells(2, 18).Value Then    'if the value in the Summary Table (row 12) is greater than the value of the Bonus cell THEN
                 ws.Cells(2, 17).Value = ws.Cells(SummaryTable, 10)     'Record Ticker ID
                 ws.Cells(2, 18).Value = ws.Cells(SummaryTable, 12)     'and Replace Greatest % Increase
                 ws.Cells(2, 18).NumberFormat = "0.00%"    'formats percentage change as a percentage
             End If
              If ws.Cells(SummaryTable, 12) < ws.Cells(3, 18).Value Then
                 ws.Cells(3, 17).Value = ws.Cells(SummaryTable, 10)     'Record Ticker ID
                 ws.Cells(3, 18).Value = ws.Cells(SummaryTable, 12)     'and Replace Greatest % Decrease
                 ws.Cells(3, 18).NumberFormat = "0.00%"    'formats percentage change as a percentage
              End If
              If ws.Cells(SummaryTable, 13) > ws.Cells(4, 18).Value Then
                 ws.Cells(4, 17).Value = ws.Cells(SummaryTable, 10)     'Record Ticker
                 ws.Cells(4, 18).Value = ws.Cells(SummaryTable, 13)     'and Replace Greatest total volume
                 ws.Columns("R").EntireColumn.AutoFit      'formats the colums to autofit to the contents for total volume
              End If
            Next SummaryTable   'Once it iterates through all of the rows with data, the Final (max or min) values will output on the spreadsheet
 '____________________________________________________________________________________________________________________________________________________________________________
'Wrap things up!

    Next ws     'move to the next worksheet
    
End Sub

