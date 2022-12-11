# An Analysis of Green Stocks

## Overview of Project

### Purpose

The purpose of this project is to refactor a VBA script that creates a report of the data of daily trade volumes and prices of 12 green stocks from the years **2017** & **2018**. The analysis highlights the following, sorted by stock ticker symbol:

* The **total daily volume** of shares/stocks traded
* The **annual return** amount

The original script looped through the entire table of stock data by an amount equal to the amount of unique tickers being analysed. The objective is to get the code to go through the stock data only one time, and to measure the runtime speeds of the old vs new code.
 
 
## Results

### Analysis of Data

#### Explanation

Using the VBA script, an analysis of the data was created. Sorted by ticker symbol, the analysis displays the **Total Daily Volumes** and **Return** amount. The total daily volume is the **sum** of every recorded daily volume for each respective stock. The return column outlines the percentage increase/decrease of stock price from the beginning of the year to the price at end of the year.


##### 2017

![green_stocks_analysis_2017](/resources/green_stocks_analysis_2017.png)


##### 2018

![green_stocks_analysis_2018](/resources/green_stocks_analysis_2018.png)


#### Brief Discussion

2017 was a great year for investing in 11 of 12 of the green stocks being analyzed. A great year in general for investing in green stocks. 

2018 was a risky year to invest in green stocks, with only 2 of 12 being analyzed seeing positive returns. The 2 stocks that had positive returns, which both had positive returns the year prior, would have made great 2 year investments.


### Refactoring of the Code

#### Explanation

The main reason for refactoring the code was due to the code originally going through a `For i = 0 To 11` loop through the entire dataset, as it worked its way through the array indexing of the **tickers**. This needed to be simplified to a single loop.

#### Approach


