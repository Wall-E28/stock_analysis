# Green Stock Analysis

## Project Overview

A friend named Steve has asked me to analysis a green stock for his parents to see if it is worth 
investing in. In order to do this, I needed to use the Visual Basic Appication in Excel to find the 
stock's total daily volume and annual return. After completing this analysis, I continued to analyze 
11 more green stocks to see how the first one compared to them. For there, I was able to use the
analysis to let Steve know what would be the best option for his parents.

### Purpose

The purpose of this project was to make an efficient way to look at multiple stocks using VBA. After 
being able to run the analysis of these 12 different stocks the first, it was apparent that there was 
a more efficient way to work with the data given. In order to make the analysis more efficient, I needed 
to refactor my code. This project looks to see if my refactoring made the analysis more efficient.

## Results

### Refactoring the Code

In order to make my code more efficient, I needed to switch the nesting order of my for loops. To do this,
I created a 4 different arrays; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 
The tickers array was used to establish the ticker symbol of a stock. I matched the other three arrays 
with the tickers array by using a variable called the tickerIndex. 

#### New Code
"""
    '3) Initialize array of all tickers
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
    
    '4a) Activate data worksheet
    Worksheets(yearValue).Activate
    
    '4b) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '5a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '5b) Create three output arrays
    
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
    '6a) Initialize ticker volumes to zero
        
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    '6b) loop over all the rows
    
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
            
            
        End If
        
        '7c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 7).Value
            

            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
"""

This variable allowed me to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to 
each ticker symbol before interating through the data set. By doing it this way, the analysis would be 
completed much faster than using the nested for loop for earlier.

###
