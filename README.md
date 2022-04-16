# Stock-analysis
## Stock Analysis with VBA
Module 2 Analysis by Lauren Debes
## Overview of Project and Purpose
Steve, a recent finance graduate, wants help researching green energy stocks to help his parents make good investment decisions. We will be using Excel VBA to analyze, make calculations, and automate data about stocks. 
### Purpose
We want to determine which green stocks have been successful (with a high yearly return) and include information on how often it is traded (daily volume). Steve's parents want to invest in DAQO New Energy Corporation (DQ), but Steve suspects it would be wiser to diversify their investments.
## Results and Analysis
### Results for Steve
The only tickers that had an increase in return in both 2017 and 2018 are ENPH and RUN. Those stocks may be wise investments for Steve's parents. They both have a high total daily volume as well. Almost all stocks (except one) had a positive rate of return in 2017, then in 2018 they almost all had a loss on return except ENPH and RUN (which suggests they may be more stable investments).
### Analysis of Code and Output
In our code we created a TickerIndex to more quickly access the index of the four arrays. In our original stock analysis, we only used one array--tickers; in refactored, we made ticker volumes, starting prices, and ending prices into arrays. This created loops that were more efficient. We used the iterator (i) to stand in for a specific index in our code:
```
tickerVolumes(i) = 0
and
Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = tickerVolumes(i)
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```
Our original code ran 2017 in 1.425781 seconds, and 2018 in 1.414063 seconds.
Our new code ran far more quickly.
![VBA_Challenge_2017.png](https://raw.githubusercontent.com/LaurenDebes/Stock-analysis/main/VBA_Challenge_2017.png)
![VBA_Challenge_2018.png](https://raw.githubusercontent.com/LaurenDebes/Stock-analysis/main/VBA_Challenge_2018.png)

## Summary
Some advantages to refactoring code include:
  * Can improve the readability of code
  * There is a functional code to fall back on if needed
  * Can improve the speed of code

Some disadvantages to refactoring code include:
  * There is a risk of messing up code that already exists, you will want to make sure the original code stays intact
  * Can be unnecessary use of time if code is already working well and there is no major benefit

Original VBA Script Pros and Cons:

| Pros  | Cons |
| ------------- | ------------- |
| Shorter code  | neglected the ticker index, which would make it faster to run |
| already fairly quick, ran in just over 1 second  | decreased speed from not using arrays |

Refactored VBA Script Pros and Cons:


| Pros  | Cons |
| ------------- | ------------- |
| Ran faster  | script was already fairly quick, may not be worth the time to refactor |
| Utilized arrays with good variable names, which made for better readability AND speed | more complex code |
