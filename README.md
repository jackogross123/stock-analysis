# Stock Analysis with VBA

## Overview of Project

### Purpose
Steve is a friend who's parents are very interested in investing in green energy. The purpose of the analysis was to determine if certain green stocks had positive or negative returns in 2017 and 2018. After I created the code I refactered it to make it cleaner and more efficient. The code that was created in the challenge was very similar for what I created in the assignment, however, with less variables and a much simpler format. Additionally, I wanted to decrease the time that it took to run the macro.

### The Data
The data for the project consisted of stock data for 12 different stocks. The information contained the tickers, the date, the high and low prices of the day, the closing price, and the daily volume. The goal of the challenge was to extract the ticker, daily volume, and annual return for each of the 12 stocks.

## Results
At first glance, the better of the two years for green stock returns was 2017. The initial goal of the analysis was to determine if the stock DQ was a solid investment for Steven's parents. After performing the analysis on DQ alone, I found that the stock had performed very poorly in 2018. I'd reccomend that Steven does not invest his parent's money in one stock, but accross a range of different tickers or an ETF. 

![DQ_Analysis](https://github.com/jackogross123/stock-analysis/blob/main/Resources/DQ_Analysis.png)

As it can be seen through the image, the 2018 return for DQ was over -62%.

With a few changes to the code I was able to change the code to be able to run an analysis on all of the tickers in 2017 or 2018. The VBA script can be found below.

Sub yearValueAnalysis()
    Dim startTime As Single
    Dim endTime As Single

yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

'1. Format the output sheet on the "All Stocks Analysis" worksheet.
Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
     'Create a header row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
    
'2. Initialize an array of all tickers.
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

'3. Prepare for the analysis of tickers.
    'Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim closingPrice As Single
    
    'Activate the data worksheet.
    Worksheets(yearValue).Activate
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4. Loop through the tickers.
For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0

'5. Loop through rows in the data.
    Worksheets(yearValue).Activate
       For j = 2 To RowCount
    
    'Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
    'Find the starting price for the current ticker.
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
        End If
    'Find the ending price for the current ticker.
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
             endingPrice = Cells(j, 6).Value
         End If
        Next j

Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
  
dataRowStart = 4
dataRowEnd = 15
For q = dataRowStart To dataRowEnd

    If Cells(q, 3) > 0 Then
        Cells(q, 3).Interior.Color = vbGreen
    ElseIf Cells(q, 3) < 0 Then
        Cells(q, 3).Interior.Color = vbRed
    Else
        Cells(q, 3).Interior.Color = xlNone
    End If

Next q

'6. Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

Next i

    endTime = Timer
    MsgBox "This Code ran in " & (endTime - startTime) & "seconds for the year " & (yearValue)

End Sub

### 2017 Analysis
In 2017, all but one stock had positive annual returns. The only stock that didn't have positive returns wsa $TERP. What's even more significant, is that over 25% of the stocks in 2017 had gains that were over 100%. From the data from the 12 tickers, it could be said that 2017 was very good year for green companies.

![VBA_Challenge_Data](https://github.com/jackogross123/stock-analysis/blob/main/Resources/VBA_Challenge_Data_2017.png)

In 2017, DQ had a remarkable performance and was up nearly 200%. SEDG also had a great run with a 184% gain.

The performance of the refactored code is noteworthy. After the initial run, the code used in the module was taking over 1 second and even 2 seconds to run in some occasions. 

![Module_2_2017](https://github.com/jackogross123/stock-analysis/blob/main/Resources/Module2_2017.png)

