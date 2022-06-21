# An Analysis of Stocks

## Overview of Project

### Purpose

The basis of this project was to refactor code originally utilized to perform an analysis on the Total Daily Volume of twelve stocks, from the years 2017 to 2018. The code was refactored to run more efficiently, accounting for the possibility of larger quantities of data needing to be analyzed with the same script. The code will also produce the total time it took to pull and analyze the selected data. The analysis was conducted to investigate the stocks' returns. The script was written in VBA to be used in Microsoft Excel.

## Results

One of the main objectives of this project was the analysis on twelve stocks from 2017 to 2018. Within the twelve stocks, we peered into the data to determine which stocks made profit, or a loss. These are indicative of positive returns, and negative returns, respectively.

![Stocks_For_2017](https://user-images.githubusercontent.com/106129195/174530886-640ad5dd-9c36-459b-a753-25b9b7657f6c.png)  ![Stocks_For_2018](https://user-images.githubusercontent.com/106129195/174531054-e02d75e0-c203-4def-8ba6-f8f55881615e.png)

###### The twelve stocks we analyzed for the years of 2017 to 2018, and their percentage of return. Green indicates profit, and red indicates loss.

Looking to the year of 2017, we can see that from the selection of stocks we analyzed, most had a positive return except for the stock **TERP**. Notably, **DQ**, **SEDG**, **ENPH** and **FSLR** did well in terms of profit. Conversely, for the year of 2018, we can see that many experienced a loss of profit. Out of our sample pool, **DQ** had taken the largest loss, whereas **ENPH** and **RUN** were the only stocks that performed well.

A key factor we also took into consideration was the run time of the script.

![VBA_Challenge_2017_OriginalScript](https://user-images.githubusercontent.com/106129195/174533741-c826d360-8925-45ec-8538-c4119dc0c79c.png)  ![VBA_Challenge_2018_OriginalScript](https://user-images.githubusercontent.com/106129195/174533749-77f13463-e8b2-4291-a287-5cbbc1e932ef.png)

###### The run time for the original script for both years.

As we can see, the run times for both years were similar, between 0.83 and 0.85 seconds. While the run time is under a second, we can refactor the code to have it run more efficiently and possibly account for larger amounts of data. Upon refactoring the code, we reviewed the run times again.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/106129195/174534242-a0d644f4-cd1b-4c71-9c26-e1eab4c7eaf5.png)  ![VBA_Challenge_2018](https://user-images.githubusercontent.com/106129195/174534255-b1597260-d4ab-4be1-85c2-69750f385b4d.png)

######  The run time for the refactored script for both years.

Comparing the run time of the refactored code to the original, we can see that the code runs faster by 0.62 and 0.61 seconds, respectively. There was a significant decrease in run time, indicating that our new refactored code does run more quickly.

This leads into our secondary objective of this project: refactoring the original VBA script. Our original script, and the refactored script, share similarities: the flow of logic of the script, and the functionality. Their differences mainly lie in how the script loops through the rows of the data as it is being ran.

Let us compare the main differences between the two scripts.

Our original VBA script for looping through the stock tickers is as follows:
```
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        '5) Loop through rows in the data
        Worksheets(yearValue).Activate
        
            For j = 2 To RowCount
            
                '5a) Get total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
            
                '5b) Get starting price for current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
        
                '5c) Get ending price for current ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
                
            Next j
```

Here is the above script snippet refactored:

```
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
        
    'Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
            'Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
        'Check if the current row is the last row with the selected ticker
        'If the next row's ticker doesn't match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If

        'Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                
                tickerIndex = tickerIndex + 1
            
            End If
    
        Next i
```
We can draw the conclusion that utilizing arrays contributed to the script running faster than the original script. The array function allows us to be more specific in the script, and group the data in a more compact form for analysis.

## Summary

Based on what has been presented, we can infer that there are advantages and disadvantages to refactoring code. One major advantage is that the code will be more precise, having fewer steps to take, and therefore execute the script at a quicker rate. This is a positive aspect to the user, as it will allow them to work more quickly and/or input a larger quantity of data to analyze. However, a disadvantage to that is that due to the code becoming more precise, therein lies the chance of running into bugs while refactoring the code. As a result, it could possibly take time to run iterations of the code, investigate the line of code that's causing the bug, and debug it.

Another advantage is that refactoring the code could make it easier for other users to read. Should the script need to be reviewed by other programmers, they could logically follow what it does and, if needed, make changes as necessary. In contrast, however, there is no unified manner of refactoring code. There is no "standard", so to speak. So while that allows more room for different routes of completing the task, it could prove difficult for the user to determine a starting point and how to go about it.

Going over the advantages and disadvantages of refactoring code, we can now compare them to our original and refactored scripts. The screenshots of the run times for the scripts show that the refactored script ran considerably faster than the original script. However, during the process of refactoring, I ran into multiple debugging issues. As the original script utilized variables, and the refactored script utilized arrays, it took me time to learn to properly utilize the array function in VBA.

Further, upon reviewing the refactored script, we can see that it is more well-defined and clear in what its objective is. Every value and array is clearly defined in what it is referencing, and the flow of the script is logical. While there was not a defined guide for how it should look, we referenced the original script to ensure a logical flow of process.
