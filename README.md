# Stock Analysis with VBA

## Overview of Project
---
To assist Steve in his stock analysis, a VBA script was built that compared the **total volume** and **return** for a given year across multiple stocks. The data itself encompassed 12 stocks, and covered the years 2017 and 2018. 

To further improve upon the code written, the script was refactored and timed for performance. Refactoring is the processes of adjusting code to make it more efficient while not affecting it's function or output. 

## Results
---
### Initial Code
First, the following was the initial code written to execute the analysis. As an immediate observation, the code was broken up into two separate sub routines.
```vbscript
Sub yearvalue()
    
    Dim yearvalue As String
    Dim startTime As Single
    Dim endTime  As Single
    
    yearvalue = InputBox("which year would you like to analyze?")
    
    startTime = Timer
    
    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" & yearvalue & ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '2) Initialize array of all tickers
    Dim tickers(11) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
   
    '3a) Initialize variables for starting price and ending price
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    '3b) Activate data worksheet
    Worksheets(yearvalue).Activate
    
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '5) loop through rows in the data
        Worksheets(yearvalue).Activate
        For j = 2 To RowCount
           
            '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value

            End If
            '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If

            '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value

            End If
        Next j
        
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
    
    'Call subroutine that applies formatting to the results
    Call formatAllStocksAnalysisTable

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub
```
```vbscript
Sub formatAllStocksAnalysisTable()

    'activate sheet to format
    Worksheets("All Stocks Analysis").Activate
    
    'formatting of header row
    'Range("A3:C3").Font.Bold = True
    'Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    With Range("A3:C3")
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Font.Size = 14
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Color = RGB(148, 33, 146)
    End With
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    datarowstart = 4
    datarowend = 15
    
    For i = datarowstart To datarowend
        
        If Cells(i, 3) > 0 Then
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then
            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
        
        Else
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If
    
    Next i


End Sub
```
### Refactored Code
---
Next, the code was refactored into the following script:

```vbscript
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    'Ask for user input to identify which year to analyze
    yearvalue = InputBox("What year would you like to run the analysis on?")

    'Begin timer to evaluate how fast the script executes.
    startTime = Timer
    
    'Activate the output sheet and apply preliminary formatting
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For r = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(r, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(r, 1).Value = tickers(tickerIndex) And Cells(r - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(r, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(r, 1).Value = tickers(tickerIndex) And Cells(r + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(r, 6).Value
        End If

            '3d Increase the tickerIndex.
        If Cells(r, 1).Value = tickers(tickerIndex) And Cells(r + 1, 1) <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    
    'Loop through output and apply conditional formatting
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'End timer and print results
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub
```
### Key Differences:

1. Testing all indexes against one row at a time **vs.** all rows against one index at a time.  
    - Original:
    ```vbscript
    'Iterate through ticker indexes
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        'loop through ALL rows for each index
        Worksheets(yearvalue).Activate
        For j = 2 To RowCount
        'From here, multiple If statements
    ```
    - Refactored:
    ```vbscript
    For r = 2 To RowCount
        'Leverage the variable tickerIndex to iterate through ticker indexes without a for loop    
        tickerVolumes(tickerIndex)....
    ```

2. Arrays: 
    - Storing data for each Ticker without exiting the `For` loop.
    - Similar to above, printing all the data to the output sheet without having to re-engage the `For` loop. 
    ```vbscript
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    ```
### Time
For the original code:

![Original: 2017](/Resources/greenstocks2017.png)![Original: 2018](/Resources/greenstocks2018.png)

For the refactored code:

![Refactored: 2017](/Resources/VBA_Challenge_2017.png)![Refactored: 2018](/Resources/VBA_Challenge_2018.png)

## Summary

### Refactoring - Pros and Cons
Refactoring code has undeniable benefits, both in terms of code performance and the readability/cleanliness of the code itself. Often times, through trial and error or simply by breaking challenges down into segments, it is understandable that one does not initially produce the most linear solution. However, once a solution is obtained, alternate paths to the same end can be explored.

In this stock analysis exercise, the benefits of refactoring can be quantified, as there was a decrease in the amount of time it took for the code to execute. While this might not be a large decrease, the time saved could add up to a significant amount if this were a subroutine in a string of many other subroutines, or this code were to be run many times in a row. 

Regarding potential disadvantages, notes concerning original work could be lost, and these notes could provide valuable insights when retracing steps and/or troubleshooting. Additionally, if the script is not to be used repeatedly or as part of a larger process, it is worth considering the cost of time spent refactoring the code against the benefit of the marginal processing power gained. 
***"If it ain't broke, don't fix it."*** - Bert Lance

### Original vs. Refactored VBA - Pros and Cons
Regarding the VBA script contained in this exercise, the implementation of arrays and the reordering of `for` loops provided a considerable benefit. Though the same output can be achieved in multiple ways, being able to store data without exiting and restarting a loop saved a noticeable amount of processing. In the latest refactored form, the script is expandable in the fact that it would continue to run efficiently even if the ticker index were made larger and/or the amount of data to analyze were to grow.

There are few disadvantages to speak of in this case. Though the original script was "simpler" in terms of fewer variables and arrays, these elements are minor in number.

