# stock-analysis
Module 2 work
# VBA Stock Analysis

## Overview of Project
The original project consisted of preparing a workbook of data analysing the performance of green energy stocks for 2017 and 2018.
The output included highlighted stock performances and allowed visual cues to easily determine poor versus successful stock prices throughout the year.
As a follow-up to the original code, refactoring was done to accomodate for a larger data set than the original smaller set of green energy stocks.
Performance of the code could easily be tested by refactoring and monitoring the time for the code to run.

## Results
In order to refactor the code multiple changes were made to the original analysis code.
First, an index was created to condense the code in the original script and three output arrays were created:
```
tickerIndex = 0

Dim tickerVolumes (12) As Long
Dim tickerStartingPrices (12) As Single
Dim tickerEndingPrices (12) As Single

```
Second, I created 2 for loops. The first to loop contains if then statements to determine first and last rows of specific tickers, with the second also containing an inrease to the tickerIndex variable:
```
If (Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex)) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

If (Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex)) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
```

Finally, the additional refactor reformatted the output to include the new output arrays completed in the earlier step:
```
Cells(k + 4, 1).Value = tickers(k)
        Cells(k + 4, 2).Value = tickerVolumes(k)
        Cells(k + 4, 3).Value = (tickerEndingPrices(k)) / (tickerStartingPrices(k)) - 1
        
    Next k
```

The below results are for pre-refactoring:

![2017](https://github.com/rbinghamjr/stock-analysis/blob/main/Pre_Refactor_Runtime_2017.PNG)

![2018](https://github.com/rbinghamjr/stock-analysis/blob/main/Pre_Refactor_Runtime_2018.PNG)

These results are for after refactoring:

![2017](https://github.com/rbinghamjr/stock-analysis/blob/main/VBA_Challenge_2017.PNG)

![2018](https://github.com/rbinghamjr/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

As shown, the refactoring decreased the amount of time to for the code to run and would significantly improve performance for a larger data set.

## Summary
In conclusion, the value of writing this code was shown to be a great way to present the data at the request of a client.
Refactoring the code showed a decrease in time for the code to run. Also, refactoring allowed the ability to condense the original code to populate the same results in fewer lines.
The time it took to refractor can negatively effect the process including having to debug the refactor when errors are found.
Overall, there is more benefit to refactoring due to the overall time save and the ability to run through a larger data set if provided.
