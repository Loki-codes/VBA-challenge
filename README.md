# VBA-challenge
Home Work week 2
I wasn't sure how to get the VBA files seperate from the excel file so I uploaded the entire excel file as xlms. I also put the vba for all of the Macros below here.

Thank-you,
Kyle


Sub plaAll()
'This macro is to play through all of the sheets with all of the macros and is attached to a button on the first sheet.
Dim wksht As Worksheet
Application.ScreenUpdating = False
For Each wksht In Worksheets
    wksht.Select
    'this calls all three macros on each worksteet as it is selected
    Call setHeaders
    Call StockData
    Call boNus
Next
Application.ScreenUpdating = True
End Sub

Sub StockData()
'I like to first define all of my variables
'this if to find the last row for the fo loop
Dim lastRow As Long
'variable to store the yearly change
Dim yearlyChange As Double
'variable to store the percent change
Dim percentChange As Double
'variable to store the volume
Dim volume As Double
'variable to move the new info down one row at a time as it creates the new table
Dim tableRow As Integer
'variable to store amount of times the else statement loops so that we can subtract that from i and end up with the first opening price of each ticker.
Dim rowCount As Double

'I then assigned all variables to what would be needed
rowCount = 0
yearlyChange = 0
percentChange = 0
volume = 0
tableRow = 2
lastRow = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row
'loop through data
For i = 2 To lastRow
    'if statement which will activate when is senses we are changing tickers.
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
        yearlyChange = Cells(i, 6) - Cells(i - rowCount, 3)
        
        'this if else statement was addedue to "0" populating into thepercentChange equasion
        If (yearlyChange = 0) Or (Cells(i - rowCount, 3) = 0) Then
            percentChange = 0
        Else
            percentChange = yearlyChange / Cells(i - rowCount, 3)
        End If
        'sum of (i, 7) to calculate Volume
        volume = volume + Cells(i, 7)
        
        'the below lines populate the variables and Ticker name into the new table
        Cells(tableRow, 9) = Cells(i, 1)
        Cells(tableRow, 10) = yearlyChange
        Cells(tableRow, 11) = percentChange
        Cells(tableRow, 12) = volume
        

        
        'formating section which covers fill colors and data types
        Cells(tableRow, 11).NumberFormat = "0.00%"
        Cells(tableRow, 10).NumberFormat = "$#,##0.00"
        If Cells(tableRow, 10) >= 0 Then
            Cells(tableRow, 10).Interior.Color = RGB(0, 255, 0)
        Else
            Cells(tableRow, 10).Interior.Color = RGB(255, 0, 0)
        End If
        If Cells(tableRow, 10) >= 0 Then
            Cells(tableRow, 11).Interior.Color = RGB(0, 255, 0)
        Else
            Cells(tableRow, 11).Interior.Color = RGB(255, 0, 0)
        End If

        'add one to the new table row so that it moves down and populates the next row
        tableRow = tableRow + 1
        'reset all variables used for previous ticker
        rowCount = 0
        yearlyChange = 0
        percentChange = 0
        volume = 0
        
    Else
    'used for when it is still the same ticker.
        'add the volume up
        volume = Cells(i, 7) + (volume)
        'increase the row count so that we can go back to the first opening price
        rowCount = rowCount + 1

    End If
Next i
End Sub
Sub setHeaders()
'
' A simple macro to fill in all of the titles that I wanted.
'

'
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Percent Change"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "Greatest % Increase"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "Greatest % Increase"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "Greatest % Decrease"
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "Greatest Total Volume"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Range("R2").Select
End Sub
Sub boNus()
'again I like defining all of my variables at the top
Dim lastRow As Double
Dim bestIncrease As Double
Dim bestDecrease As Double
Dim greatestVolume As Double
Dim bestIncreaseTkr As String
Dim bestDecreaseTkr As String
Dim greatestVolumeTkr As String

'find last row of each sheet
lastRow = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row
'set variables to 0
bestIncrease = 0
bestDecrease = 0
greatestVolume = 0
For i = 2 To lastRow
    'if statement that looks for a value higher than the previous highest value
    If Cells(i, 11).Value > bestIncrease Then
        'resets the variable to the new highest
        bestIncrease = Cells(i, 11).Value
        'records current ticker
        bestIncreaseTkr = Cells(i, 9).Value
        'populates new table with collected data
        Cells(2, 17) = bestIncreaseTkr
        Cells(2, 18) = bestIncrease
        'formatting for data in table
        Cells(2, 18).NumberFormat = "0.00%"
        Cells(2, 18).Interior.Color = RGB(0, 255, 0)
    End If
    
    'if statement that looks for a value lower than the previous lowest value
    If Cells(i, 11).Value < bestDecrease Then
        bestDecrease = Cells(i, 11).Value
        'records current ticker
        bestDecreaseTkr = Cells(i, 9).Value
        'populates new table with collected data
        Cells(3, 17) = bestDecreaseTkr
        Cells(3, 18) = bestDecrease
    End If
     
     'if statement that looks for a value higher than the previous highest value
     If Cells(i, 12).Value > greatestVolume Then
        'resets the variable to the new highest
        greatestVolume = Cells(i, 12).Value
        'records current ticker
        greatestVolumeTkr = Cells(i, 9).Value
        'populates new table with collected data
        Cells(4, 17) = greatestVolumeTkr
        Cells(4, 18) = greatestVolume
        'formatting for data in table
        Cells(4, 18).NumberFormat = "0.00E+00"
    End If
    
    
Next i
End Sub
