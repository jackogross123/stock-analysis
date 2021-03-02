# Stock Analysis with VBA
The excel workbook can be found here: [VBA_Challenge](https://github.com/jackogross123/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project

### Purpose
Steve is a friend who's parents are very interested in investing in green energy. The purpose of the analysis was to determine if certain green stocks had positive or negative returns in 2017 and 2018. After I created the code I refactered it to make it cleaner and more efficient. The code that was created in the challenge was very similar for what I created in the assignment, however, with less variables and a much simpler format. Additionally, I wanted to decrease the time that it took to run the macro.

### The Data
The data for the project consisted of stock data for 12 different stocks. The information contained the tickers, the date, the high and low prices of the day, the closing price, and the daily volume. The goal of the challenge was to extract the ticker, daily volume, and annual return for each of the 12 stocks.

## Results
At first glance, the better of the two years for green stock returns was 2017. The initial goal of the analysis was to determine if the stock DQ was a solid investment for Steven's parents. After performing the analysis on DQ alone, I found that the stock had performed very poorly in 2018. I'd reccomend that Steven does not invest his parent's money in one stock, but accross a range of different tickers or an ETF. 

![DQ_Analysis](https://github.com/jackogross123/stock-analysis/blob/main/Resources/DQ_Analysis.png)

As it can be seen through the image, the 2018 return for DQ was over -62%.

With a few changes to the code I was able to change the code to be able to run an analysis on all of the tickers in 2017 or 2018. 

### The Refactored Code
    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

### 2017 Analysis
In 2017, all but one stock had positive annual returns. The only stock that didn't have positive returns wsa $TERP. What's even more significant, is that over 25% of the stocks in 2017 had gains that were over 100%. From the data from the 12 tickers, it could be said that 2017 was very good year for green companies.

![VBA_Challenge_Data](https://github.com/jackogross123/stock-analysis/blob/main/Resources/VBA_Challenge_Data_2017.png)

In 2017, DQ had a remarkable performance and was up nearly 200%. SEDG also had a great run with a 184% gain.

The performance of the refactored code is noteworthy. After the initial run, the code used in the module was taking over 1 second and even 2 seconds to run in some occasions. 

![Module_2_2017](https://github.com/jackogross123/stock-analysis/blob/main/Resources/Module2_2017.png)

After I refactored the script the run time decreased a significant amount.

![VBA_Challenge_2017](https://github.com/jackogross123/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

The time to run the script after I refactored it was down to nearly a millisecond.

### 2018 Analysis
Unfortunately for green investors, 2018 wasn't the best year for green stocks. Companies such as DQ and SEDG, which had incredible runs in 2017, retracted a little from their yearly opens and finished the year in the red.

![VBA_Challenge_Data_2018](https://github.com/jackogross123/stock-analysis/blob/main/Resources/VBA_Challenge_Data_2018.png)

Only 2 of the 12 tickers had posititve returns. ENPH and RUN both had high returns in 2018. ENPH's 2018 performance is notable because in 2017 ENPH was up nearly 130%. For fun, I wanted to take a look at DQ's performance in 2020. I know that in 2020, green stocks performed particularly well. As an investor in green energy myself, many of my clean energy holdings have performed well. DQ's return in 2020 was over 1000%. This is due to the rally behind solar for envionrmental reasons, China's dominance over PV panel production, and the effects of the pandemic.

As seen in the 2017 analysis, the refactored code improved the time it took to run the script.

![Module_2_2018](https://github.com/jackogross123/stock-analysis/blob/main/Resources/Module2_2018.png)

It took over 2 seconds to run the all years analysis from the Module for year 2018. After refactoring the script, the run time was decreased greatly.

![VBA_Challenge_2018](https://github.com/jackogross123/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The time it took to run the code was cut down to nearly 0.20 seconds.

## Summary

### Advantages and Disadvantages of Refactoring Code
I think that one of the biggest advantages to refactoring code is just making it more readable and efficient. With my little experience so far, I have found that I often get confused with what I'm working on, especially if there are lots of different lines of code. I think that the best strategy moving forward is to create a really good outline before I start to working on the code. I have to ensure that every step is accounted for so I can stay organized through the refactoring process. Additionally, refactoring the code forces you to organize it which makes it much easier to read for an outside audience.

I think that one of the biggest disadvantages of refactoring code is possibly breaking the code. Obviously this can be prevented by saving the file frequently, but for someone who still has weak VBA knowledge, it can be a little intimidating to look at a script and find ways to improve it and organize it without truly understanding the purpose of every line.

### Advantages and Disadvantages of Refactoring the Code from Module 2
One of the biggest advantages from refactoring this specific code was organizing it which in turn helped me understand many of the parts that I didn't understand before. I am still struggling a little bit with for loops, but the challenge helped a lot with that. Another advantage to refactoring this code was improving the time it took to run the anaylsis. The code from the module was clunky and wasn't very efficient, so by refactoring it we solved both of those problems. 

I think that one of the biggest disadvantages and challenges for me was working with the new loops and the index. I didn't understand a lot about the tickerIndex and I'm still struggling to find out why an index was needed in order to run the script. Additionally, when I first tried to run the script I kept getting an "overflow" error message from Excel. This was because I had too many loops and "next i" lines which prevented the code from being ran. This is something that I'm still working on, but I know that with enough practice I'll soon start to understand all of these things. Overall, it can be tough to try and improve something when you don't necessarily understand the purpose of every part of the mechanism. 
