# Stock Market Analysis

## Overview of Project
This project conatins VBA subroutines to calculate the total daily volume and yearly return
from the 2017 and 2018 data for the 12 stocks in `VBA_Challenge.xlsm`. The
total daily volume is the sum of the total number of shares traded each day and
then totaled over the entire year and therefore shows the relative activity of
each stock, while the yearly return shows the percent change in the stock price
from the beginning to the end of the year. The subroutines `AllStocksAnalysis`
and [`AllStocksAnalysisRefactored`](VBA_Challenge.vbs) calculate these
quantities for each stock in sheets `2017` and `2018` and then display the
results in sheet `All Stocks Analysis`.

### Purpose
The purpose of this project is to compare execution times for the subroutines
`AllStocksAnalyis` and [`AllStocksAnalysisRefactored`](VBA_Challenge.vbs) to
explore the advantages and disadvantages of refactoring code and its
application to this analysis.

## Results
### `AllStocksAnalysis`
The first version of this subroutine `AllStocksAnalysis` (workbook module 1)
loops through each row for either 2017 or 2018 of `VBA_Challenge.xlsm` for each
ticker contained in the dataset. It does so through the following nested loop:
```
For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        'Find total volume for the current ticker
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If

        'Find the starting price for the current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        End If

        'Find the ending price for current ticker
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            endingPrice = Cells(j, 6).Value
        End If
    Next j
    'Output results for the current ticker
```
It is noted the we can acquire the starting and ending prices by checking for a
change in the current ticker from the previous and next rows because the dataset
is grouped by ticker and then ordered chronologically. We then calculate the yearly
return from the starting and ending prices and output the results to
`All Stocks Analysis`. This executes in roughly 0.6 seconds for either year.

### `AllStocksAnalysisRefactored`
In the refactored subroutine
[`AllStocksAnalysisRefactored`](VBA_Challenge.vbs) (workbook module 2), we
replace the nested loop with a single loop and acquire the total volume and
yearly return for each stock in a single iteration through the dataset. We thus
use arrays `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` to
store the total volume and starting/ending prices for each stock along with the
index variable `tickerIndex` to keep track of the current ticker as we loop
through the rows of either year's data. Our loop then becomes:
```
For i = 2 To RowCount
    'Increase volume for the current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    'Check if current row is the first row with the selected tickerIndex
    'If so, note the starting price
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If

    'Check if current row is the last row with the selected tickerIndex
    'If so, note the ending price and increase tickerIndex
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        tickerIndex = tickerIndex + 1
    End If
Next i
'Output results for all tickers
```
This is advantageous in terms of computation time due to the removal of the
nested loop. The previous version required an iteration through the entire
dataset for each of the 12 tickers. This is unnecessary since the data is
ordered chronologically and grouped by ticker, and so most iterations in the
inner loop result in no computation due to the conditionals. The new version
removes these unnecessary iterations to significantly decrease the execution
time. Further changes include reordering to open the dataset first and acquire
the quantities of interest and then open the output sheet to display the
results instead of switching between the two, removing repeated code. These
changes reduce the execution time to 0.09 seconds for
[2017](Resources/VBA_Challenge_2017.png) and 0.08 seconds for
[2018](Resources/VBA_Challenge_2018.png), a decrease by roughly a factor of
six.

## Summary
### Disadvantages
We see that refactoring this program has changed its structure
and resulting execution time. Refactoring in general can be disadvantageous
as it often involves shortening procedures which can decrease readability.
`AllStocksAnalysisRefactored` exemplifies this as the reordering to remove
duplicated code makes the program flow somewhat less clear. Further, refactors
in general can have unexpected effects on previously working scripts. This
could occur here as the changes involve modifying when the relevant worksheet
is activated and thus if implemented wrong could result in overwriting data.

### Advantages
It can be argued however that the advantages of refactoring this subroutine
outweigh its disadvantages. In terms of readability, it is much easier
to understand the single loop than the nested version. Further, the removal of
duplicated code is in general advantageous as it makes later changes simpler
since we need to modify code in less places. An example here is if we change
the name of the output sheet `All Stocks Analysis`, in which case the updates
allow us to make this change in one line as opposed to two. In addition, unit
testing and version control can mitigate unexpected effects of refactoring.
Finally, the decrease in execution time is especially advantageous. This is a
common effect of refactoring as we revise pieces of code to achieve the same
result in less steps. This is difficult to achieve when formulating an
algorithm for the first time, but is much easier with the initial conditions
and expected output already in place. For this analysis, the decrease in
execution time makes the refactored subroutine more general to larger datasets
with either more stocks for wider analysis and/or a larger collection
of years to analyze changes over time.
