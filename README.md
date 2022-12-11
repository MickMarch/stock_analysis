# An Analysis of Green Stocks

## Overview of Project

### Purpose

The purpose of this project is to ***refactor a VBA script*** that creates a report of the data of daily trade volumes and prices of 12 green stocks from the years **2017** & **2018**. The analysis highlights the following, sorted by stock ticker symbol:

* The **total daily volume** of shares/stocks traded
* The **annual return** amount

The original script looped through the entire table of stock data by an amount equal to the amount of unique tickers being analyzed. The objective is to get the code to go through the stock data only one time, and to measure the runtime speeds of the old vs new code.

### Important Note for Consideration

The dataset being used is sorted first by the **Ticker** column in **Alphabetical** order, followed by a secondary sort of the **Date** column in **Ascending** order. The code relies on this format.
 
 
## Brief Analysis of Data

### Explanation

Using the VBA script, an analysis of the data was created. Sorted by ticker symbol, the analysis displays the **Total Daily Volumes** and **Return** amount. The total daily volume is the **sum** of every recorded daily volume for each respective stock. The return column outlines the percentage increase/decrease of stock price from the beginning of the year to the price at end of the year (with Green being a positive return, and Red being a negative return)


#### 2017

![green_stocks_analysis_2017](/resources/green_stocks_analysis_2017.png)


#### 2018

![green_stocks_analysis_2018](/resources/green_stocks_analysis_2018.png)


## Refactoring of the Code

### Explanation

The main reason for refactoring the code was due to the code originally going through the entire dataset **12 times**. I needed to refine the code to only going through the list **once**.

The culprit was this initial loop and its application:
`For i = 0 To 11`

The loop was passing through the entire dataset, with a function that was only concerned with the ticker in the tickers array at `tickers(i)`

### Solution

#### New Index Variable

I needed something new to replace the ***tickers*** array indexing instead of a For loop of 0 To 11:
`tickerIndex = 0`

Now I had a new variable for moving through the existing arrays of:
```
tickers(tickerIndex)
tickerVolumes(tickerIndex)
tickerStartingPrice(tickerIndex)
tickerEndingPrice(tickerIndex)
```

#### Replacing the For Loop

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

## Runtime Results

### Runtime of Old Code

#### 2017
![Before_Refactoring_2017](/resources/Before_Refactoring_2017.png)

#### 2018
![Before_Refactoring_2018](/resources/Before_Refactoring_2018.png)


### Runtime of New Code

#### 2017
![VBA_Challenge_2017](/resources/VBA_Challenge_2017.png) 

#### 2018
![VBA_Challenge_2018](/resources/VBA_Challenge_2018.png)


## Summary

### Advantages of Refactoring Code in General

1. Runtime improvements. Being able to complete tasks faster is a desirable outcome of refactoring code.
2. Memory management would be something to consider when you go through large datasets or complex tasks. This would always be an important thing to consider when refactoring.
3. An advantage that is outside of the code itself: personal growth. The experience you gain from refactoring code into a more efficient script really increases one's problem solving skills.


### Disadvantages of Refactoring Code in General

1. It can be easy to make a mistake. Sometimes it may seem like an easy take to simplify code, but it could easily introduces errors. Save your work. 
2. I imagine that the more a code is refactored into a tight and simplified solution, the more an explanation is required. Algorithms aren't known for being intuitive. I imagine the more a code is refined, more lines of comments would be required to explain the code.


### Advantages of New Code from this VBA Script

As can be seen from the runtime results from above, there is a clear runtime improvement by:

* 2017 runtime improvement = (0.2734375 / 0.0859375 -1 ) x 100 = 218.18% faster
* 2018 runtime improvement = (0.2734375 / 0.0859375 -1 ) x 100 = 218.18% faster
    
### Disadvantages of New Code from this VBA Script

1. It can be harder to read than the older code. Although it is a simple algorithm, it is more complex than the older code and could take some more explaining to someone not familiar with algorithms.
2. Limits of the code:
    * There are a lot of hard coded values that don't allow for a dataset with more than 12 predetermined ticker symbols
    * The code relies on a dataset already sorted by Ticker (Alphabetical), Date (Ascending) to be functional.

### Advantages of Old Code from this VBA Script

I only really have one point for this: the code is easy to follow. It is easier to explain to someone who is new to coding.

### Disadvantages of Old Code from this VBA Script

1. Considerably slower (as mentioned above)
2. Same limits as the new code mentioned above:
    * There are a lot of hard coded values that don't allow for a dataset with more than 12 predetermined ticker symbols
    * The code relies on a dataset already sorted by Ticker (Alphabetical), Date (Ascending) to be functional.

## Resources

The data used in this project can be found here: [VBA_Challenge](VBA_Challenge.xlsm)
