# Stock Price Analysis

## Overview of Project

### Purpose
The purpose of this project was to help Steve prepare a stock analysis for his parents. After preparing an analysis for 12 stock tickers, Steve wanted to expand the dataset to include the entire stock market over the last few years. I helped him do this by refactoring the original code to work for the entire stock market, which also helped the code run faster.

## Results
In order to expand the analysis to include the entire stock market I needed to create a tickerIndex to access the correct index over four different arrays. For example, when increasing the volume for a ticker, instead of using `tickerVolumes(i)` over the loop, I used `tickerVolumes(tickerIndex)`. I used this same index when finding the starting prices and ending prices. Using an `If` statement to determine if the current row was the first or last, I used `tickerStartingPrices(tickerIndex)` to access the correct ticker index. In order to show how fast the code ran, I included a message box that calculated the start and end time of the code. In the original code, the run time was around .6 seconds. As you can see in [Run Time 2017](Resources/VBA_Challenge_2017.png) and [Run Time 2018](Resources/VBA_Challenge_2018.png), the refactored code is running considerably faster than the original times. Based on the analysis, the two stocks that had positive returns for both 2017 and 2018 were **ENPH** and **RUN**.

## Summary

### Refactoring Code
One of the advangtages of refactoring code is that it cleans up our original code to run faster. It also makes it a lot easier to go into the code and update it. A disadvantage though is that it's time consuming to change your code.
As you can see from the results, refactoring this VBA script made it run about .4 sedonds faster. Refactoring it took some time though. In the orginial script, there is only one array, with 3 variables for volume, starting price, and ending price. It was time consuming for me to conceptualize using 3 arrays with a variable `tickerIndex` that navigated the arrays.
