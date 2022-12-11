# An Analysis of Green Stocks

## Overview of Project

### Purpose

The purpose of this project is to ***refactor a VBA script*** that creates a report of the data of daily trade volumes and prices of 12 green stocks from the years **2017** & **2018**. The analysis highlights the following, sorted by stock ticker symbol:

* The **total daily volume** of shares/stocks traded
* The **annual return** amount

The original script looped through the entire table of stock data by an amount equal to the amount of unique tickers being analyzed. The objective is to get the code to go through the stock data only one time, and to measure the runtime speeds of the old vs new code.

### Important Note for Consideration

The dataset being used is sorted first by the **Ticker** column in **Alphabetical** order, followed by a secondary sort of the **Date** column in **Ascending** order. The code relies on this format.
 
 
## Results

### Brief Analysis of Data

#### Explanation

Using the VBA script, an analysis of the data was created. Sorted by ticker symbol, the analysis displays the **Total Daily Volumes** and **Return** amount. The total daily volume is the **sum** of every recorded daily volume for each respective stock. The return column outlines the percentage increase/decrease of stock price from the beginning of the year to the price at end of the year.


##### 2017

![green_stocks_analysis_2017](/resources/green_stocks_analysis_2017.png)


##### 2018

![green_stocks_analysis_2018](/resources/green_stocks_analysis_2018.png)


### Refactoring of the Code

#### Explanation

The main reason for refactoring the code was due to the code originally going through the entire dataset **12 times**. I needed to refine the code to only going through the list **once**.

The culprit was this initial loop and its application:
`For i = 0 To 11`

The loop was passing through the entire dataset, with a function that was only concerned with the ticker in the tickers array at `tickers(i)`

#### Solution

##### New Index Variable

I needed something new to replace the ***tickers*** array indexing instead of a For loop of 0 To 11:
`tickerIndex = 0`

Now I had a new variable for moving through the existing arrays of:
```
tickers(tickerIndex)
tickerVolumes(tickerIndex)
tickerStartingPrice(tickerIndex)
tickerEndingPrice(tickerIndex)
```

##### Replacing the For Loop

The next step was to increase `tickerIndex` as it went through the dataset. Luckily at the end of the row iteration, I already had a conditional in place that detected the last entry for `tickers(tickerIndex)`, which was previously `tickers(i)`:
```
If tickers(tickerIndex) <> Cells(i + 1, 1) Then
            
    tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            
End If
```

This checks to see if the following row's ticker cell differs from the current row's ticker cell. This would mean that the next row is the start of a new stock, and would therefore be a great spot to increase the `tickerIndex`:
```
If tickers(tickerIndex) <> Cells(i + 1, 1) Then
            
    tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
    tickerIndex = tickerIndex + 1
            
End If
```

---

#### Results

##### Runtime of Old Code

![Before_Refactoring_2017](/resources/Before_Refactoring_2017.png = 200x200) ![Before_Refactoring_2018](/resources/Before_Refactoring_2018.png)

---

##### Runtime of New Code

![VBA_Challenge_2017](/resources/VBA_Challenge_2017.png) ![VBA_Challenge_2018](/resources/VBA_Challenge_2018.png)


