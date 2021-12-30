# VBA Challenge
The purpose of this analysis was to improve upon the code previously written to analyze a sampling of stocks. Once improved performance is achieved, we can leverage this same code to analyze larger datasets with less CPU consumption and latency. 

## Overview
Two different approaches were leveraged to analyze stock performance for 12 different stocks for the years 2017 and 2018. In the second approach, VBA code was refactored to more efficiently deliver results and improve the output time.

## Results

Refactoring the code resulted in the same results, with greatly improved performance times. Compare screenshots of the original code versus the refactored code. The 2017 analysis took 0.9726563 seconds to output results. When that code was refactored, our new time was 0.1875 seconds. 

### Original results from green_stocks file:
![image](https://user-images.githubusercontent.com/95661802/147776157-60998238-6050-49be-91a0-228e5aa4dc2f.png)

### Improved results from VBA-Challenge file:
![2017Screenshot](https://user-images.githubusercontent.com/95661802/147776113-8c776ad3-50cd-48ee-85f4-a26a2e11733f.JPG)

These results were achieved by a few notable improvements:
* Running separate loops instead of nested loops, which cut down on processing time. 
* Defining the StartingPrices and EndingPrices As Single, as opposed to As Long.

### Original code from green_stocks file, showing As Double and nested loops:

```
'Initialize variables for starting price and ending price
Dim startingPrice As Double
Dim endingPrice As Double

'Activate data worksheet
Worksheets(yearValue).Activate

'Get the number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Loop through tickers
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
```

### Refactored code from VBA-Challenge file, showing As Single and separate loops:

```
'1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i

        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
   Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
```
    
## Summary
The refactored code will be far superior for analyzing a larger dataset, should we decide to evaluate performance for the entire stock market. 

### Advantages or disadvantages of refactoring code
There are always alternate ways to achieve the same results when writing VBA code. Performance and simplicity should always be considered, and improved upon where possible. It is also worth noting that the added commentary to the VBA code, while having no impact on performance, can also make the code easier to read and simpler to edit, should later amendments need to be made to the code itself.

### How do these pros and cons apply to refactoring the original VBA script?

Refactoring the original script was worthwhile and created a code that can be leveraged for larger projects in the future.
