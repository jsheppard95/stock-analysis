# Stock Analysis

## Overview of Project
Here we use VBA tools to calculate the total daily volume and yearly return for
the 12 stocks in `VBA_Challenge.xlsm`. The total daily volume is the sum of the
total number of shares traded each day and then totaled over the entire year
and therefore shows the relative activity of a each stock, while the yearly
return shows the percent change in the stock price from the beginning to the
end of the year. The subroutines `AllStocksAnalysis` and
[`AllStocksAnalysisRefactored`](VBA_Challenge.vbs) calculate these quantities
for each stock in sheets `2017` and `2018` and then displays the results in
sheet `All Stocks Analysis`.

### Purpose
The purpose of this project is to compare execution times for the subroutines
`AllStocksAnalyis` and [`AllStocksAnalysisRefactored`](VBA_Challenge.vbs) to
explore the advantages and disadvantages of refactoring code and its
application to stock market analysis.

## Results
The original version of this subroutine `AllStocksAnalysis` loops through each
row for either 2017 or 2018 of `VBA_Challenge.xlsm` for each ticker contained
in the dataset. It does so through the following nested `For` loop:
```
Dim startingPrice As Single
Dim endingPrice as Single

Worsheets(yearValue).Activate
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
It is noted the we can aquire the starting and ending prices in this way
because the data is ordered chronologically and grouped by ticker. We then
calculate the yearly return from the starting and ending prices and output the
results to `All Stocks Analysis`. This executes in roughly 0.6 seconds for
either year.

In the refactored subroutine [`AllStocksAnalysisRefored`](VBA_Challenge.vbs),
we replace the nested `For` loop with a single `For` loop to aquire the total
volume and yearly return for each stock in a single iteration through the
dataset. To do so we need arrays to store the total volume and starting/ending
prices for each stock along with the index variable `tickerIndex` to keep track
of which ticker we are on as we loop through the rows of either year's data.
Our loop then becomes
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
```
This is advantageous in terms of computation time primarlily due to the
removal of the nested `For` loop. The previous version required an iteration
through the entire dataset for each of the 12 tickers. This wastes time since
we know the data is ordered chronologically and grouped by ticker, meaning the
vast majority of steps in the inner loop result in no computation from the
conditionals that check if the row applies to the current ticker. The new
version instead aquires the total volume and yearly return of each stock in a
single iteration and thus removes unnecessary computations. Further changes
include reordering to open the dataset first and to aquire the quantities of
interest and then open the output sheet to display the results as opposed to
switching between the two, removes unnessary repeated code. These changes
reduce the execution time to 0.09 seconds for
[2017](Resources/VBA_Challenge_2017.png) and 0.08 seconds for
[2018](Resources/VBA_Challenge_2018.png), a decrease by roughly a factor of
six.

## Summary
We see that refactoring this program has signficantly changed its structure
and resulting execution times. Refactoring in general can be disadvantageous
as it often involves shortenning procedures which can decrease readability.
`AllStocksAnalysisRefactored` examplifies this as the reordering to remove
duplicated code makes the analysis flow somewhat less clear. Further,
refactors in general have unexpected effects on previously working scripts.
This could occur here as the changes involve modifying when the relevant
worksheet is activated and thus if done incorrectly could result in
overwritting data.

It can be argued however that the advantageous of refactoring this subroutine
outweigh its disadvantages. In terms of readability, from a glance it is much
easier to understand the single `For` loop than the nested version. Further,
the removal of duplicated code is in general advatageous as it makes later
changes simpler since we need to modify code in less places. An example here
could be if we changed the name of the output sheet `All Stocks Analysis`, 
where the updates allow us to make this change in one line as opposed to two.
In addition, unit testing with a large coverage can mitigate unexpected
effects of refactoring. Finally, the decrease in execution time is expecially
advantageous. This is a common affect of refactoring as we revise pieces of
code to achieve the same result in less steps. This is difficult to do when
formulating an algorithm for the first time, but is far simpler when the
initial conditions and expected output are already in place. In the case of
this analysis, the factor of 6 decrease in execution time makes the refactored
subroutine more generalizeable to larger datasets with either more stocks for
a more varied analysis or a larger collection of years to analyze changes over
larger timescales.
