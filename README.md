# Written Analysis: Green Stocks

## Overview of the Project

This project's goal is to automate the steps to analyze 12 stock tickers and produce a summary report for years 2017 and 2018, enabling the user to choose which year to analyze.
The second goal of this project is to demontsrate how using array instead of nested loops can help improve the performance of a piece of code.

## Results

The VBA code that is the result of this project can be utilized to automate the analysis of any stock data, with minimal efforts.
Using VBA code, the following stock performance summary was releaved for years 2017 and 2018.

---
Y 2017
---
![image y2017](https://github.com/githubteodora/stocks-analysis/blob/main/VBA_Challenge_2017.PNG)
---
Y 2018
---
![image y2018](https://github.com/githubteodora/stocks-analysis/blob/main/VBA_Challenge_2018.PNG)
---

It turned out that for years 2017 and 2018, only 2 stocks had positive return: ENPH and RUN;

### Original code
Below is the code which uses 2 nested loops to populate the summary stats per ticker.
  Runtimes
  -  y2018: 0.78515625
  -  y2017: 0.7421875

```
Sub yearValueAnalysis()

Dim startTime As Single
Dim endTime As Single

'adding an input box where the user will emter the year; this will ne stored as a variab;e
yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer

   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   'THE 2 NESTED LOOPS FOLLOW BELOW
   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
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
   'END OF THE NESTED LOOPS

'Formatting
Worksheets("All Stocks Analysis").Activate
Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

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

endTime = Timer

If yearValue = 2018 Then
    Cells(18, 1).Value = yearValue
    Cells(18, 2).Value = endTime - startTime
ElseIf yearValue <> 2018 Then
    Cells(19, 1).Value = yearValue
    Cells(19, 2).Value = endTime - startTime
End If

'the message that will display hpw much time it took the code to run
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```

This initial version of the code worked reasonably well on the small dataset.

### Improved code

It was possible to further improve the performance of the code, so that it can be used on large datasets, too.
The improvement that was attenpted, was based on using ARRAYS instead of nested loops and is available for a review below.

Runtimes are up to 4 times faster than the previous version of the code.
 -  y2018: 0.171875 
 -  y2017: 0.171875

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
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
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'1a) Create a ticker Index variable
Dim tickerIndex As Single
tickerIndex = 0

'HERE ARE THE 3 ARRAYS THAT WERE DEFINED; THEY WILL STORE THE SUMMARY VALUES PER TICKER, USING THE TICKER INDEXES - 12 ALTOGETHER
'1b) Create three output arrays; note to self - arrays must always have brackets indicating the number of buckets they carry;
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

'THIS CODE MAKES SURE THE SUMMARY VALUES ARE ZEROED OUT BEFORE ANY NUMBERS GET STORED IN THE ARRAYS;
''2a) Create a for loop to initialize the tickerVolumes, tickerStartingPrices and tickerEndingPrices to zero.

For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i

''2b) Loop over all the rows in the spreadsheet.

For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

     'If the next row’s ticker doesn’t match, increase the tickerIndex
        '3d Increase the tickerIndex.
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next i

'HERE FOLLOWS THE CODE THAT TAKES THE SUMMARY VALUES FROM THE ARRAY AND DISPLAYS THEN FOR EACH TICKER
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
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

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

## Summary:

### Advantages and Disadvantages of Refractoring Code:

#### Advantages: 
 - code gets optimized, bugs are detected
 - code performs better
 - additional explanatory notes are added to make the code easier to read and understand

#### Disadvantages:
 - it is very timeconsuming to refractor code; it is impossible to estimate the effort from start to finish
 - it can be tricky to refractor code created by somebody else
 - the refractored code can run faster, but it can also be more difficult to understand

When it comes to the original VBA script and its refractoring - it took the author 4+ hours to complete the refractoring, debug and achieve the same summaries as with the original code.
Some of the steps required troubleshooting, theory revision (arrays) and referencing previous versions of the code. 

