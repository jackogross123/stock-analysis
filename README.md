# Stock Analysis with VBA

## Overview of Project

### Purpose
Steve is a friend who's parents are very interested in investing in green energy. The purpose of the analysis was to determine if certain green stocks had positive or negative returns in 2017 and 2018. After I created the code I refactered it to make it cleaner and more efficient. The code that was created in the challenge was very similar for what I created in the assignment, however, with less variables and a much simpler format. Additionally, I wanted to decrease the time that it took to run the macro.

### The Data
The data for the project consisted of stock data for 12 different stocks. The information contained the tickers, the date, the high and low prices of the day, the closing price, and the daily volume. The goal of the challenge was to extract the ticker, daily volume, and annual return for each of the 12 stocks.

## Results
At first glance, the better of the two years for green stock returns was 2017. The initial goal of the analysis was to determine if the stock DQ was a solid investment for Steven's parents. After performing the analysis on DQ alone, I found that the stock had performed very poorly in 2018. I'd reccomend that Steven does not invest his parent's money in one stock, but accross a range of different tickers or an ETF. 

![DQ_Analysis](https://github.com/jackogross123/stock-analysis/blob/main/Resources/DQ_Analysis.png)

The performance of the refactored code is noteworthy. After the initial run, the code used in the module was taking over 1 second and even 2 seconds to run in some occasions. 

![

### 2017 Analysis
In 2017, all but one stock had positive annual returns. The only stock that didn't have positive returns wsa $TERP. What's even more significant, is that over 25% of the stocks in 2017 had gains that were over 100%. From the data from the 12 tickers, it could be said that 2017 was very good year for green companies. 
