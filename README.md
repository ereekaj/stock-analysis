# Stock-analysis Challenge

## Overview of Project: 
The purpose of this project is to improve the workbook created for my client, Steve. The original workbook used VBA macros to help Steve analyze 12 stocks. Now he would like to analyze much more data on the entire stock market. With a lot of data, the workbook may run too slowly, so I plan to refactor or edit the code to see if there is a way to do the same analysis but use code that will be more efficient in analyzing data in general.  

## Results: 

### Stock performance 2017 vs 2018
Using both sets of code, analysis showed that stocks in 2017 had more positive returns than 2018. In 2017, 11 of the 12 stocks had positive returns.  In 2018, only 2 stocks had positive returns.  Some additional analysis on a larger dataset would be needed to determine if this was a theme across most stocks in those years or not.  
### Original vs Refactored code
Below is a screenshot of the popup recording the execution time of refactored code on 2017 data. 
![2017-screenshot](https://github.com/ereekaj/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

Below is a screenshot of the popup recording the execution time of refactored code on 2018 data. 
![2018-screenshot](https://github.com/ereekaj/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The execution time of the refactored code was much faster than the original code. The refactored code decreased the execution time for each year by over a half of a second.  Below is a graph showing the stark difference of the execution times of each of the different codes in 2017 and 2018.  The refactored code decresed the execution time by 79% on the 2017 data and by 82% on the 2018 data.  This could be very useful when analyzing larger datasets. 
![graph-time-comparison.png](https://github.com/ereekaj/stock-analysis/blob/main/Resources/Graph-time-comparison.png)

The efficiency of the refactored code can be attributed to the use of arrays which can store more than one value instead of using individual variables to store each value.  You only have to define the array once and then use an index to tell the information apart.  For example, the following code was used to establish arrays in the refactored code using (12) as the index to identify the size of the array:
```
   '1b) Create three output arrays
   Dim tickerVolumes(12) As Long
   Dim tickerStartingPrices(12) As Single
   Dim tickerEndingPrices(12) As Single
```
Although this added additional lines of code, it still made the workbook more efficient. That's because all of the stock information is stored in one place, the array and is then easier to retrieve when doing the analysis. 

## Summary: 
In a summary statement, address the following questions.
- What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code is that you can use less code that will run faster when analyzing a lot of data.  The disadvantages can be the time it actually takes to refactor the code.  I ran into a lot of problems and it may have saved time to just use the same code. 

- How do these pros and cons apply to refactoring the original VBA script?
Steve wants to analyze the entire stock market over a few years.  According to [Wikipedia](https://en.wikipedia.org/wiki/New_York_Stock_Exchange), there are 2400 stocks on the New York Stock Exchange (NYSE). The refactored code improved the time to analyze the data for 12 stocks over 2 years over a half a second.  If you apply that to the entire NYSE over the course of many years, the magnitude of of efficiency could be tremendous.  Using refactored code will help Steve save time and frustration if he doesn't mind waiting for a non-expert like me to create the code. 
