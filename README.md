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

#### Refactored Code

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
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i

#### Original Code 

    2) Initialize array of all tickers
   
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

   '3a) Initialize variables for starting price and ending price

Dim startingPrice As Double
Dim endingPrice As Double

   '3b) Activate data worksheet
   
Worksheets(yearValue).Activate

   '3c) Get the number of rows to loop over
   
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   
For i = 0 To 11
    ticker = tickers(i)
    TotalVolume = 0
    Worksheets(yearValue).Activate
    
       '5) loop through rows in the data
       
    For j = 2 To RowCount
    
           '5a) Get total volume for current ticker

     If Cells(j, 2).Value = ticker Then

            'increase totalVolume by the value in the current row
            TotalVolume = TotalVolume + Cells(j, 9).Value
    
    End If
    
           '5b) get starting price for current ticker

        If Cells(j - 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set starting price
            startingPrice = Cells(j, 7).Value

        End If

           '5c) get ending price for current ticker
           
           If Cells(j + 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set ending price
            endingPrice = Cells(j, 7).Value

        End If

       Next j
       '6) Output data for current ticker

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = TotalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i



This variable allowed me to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to 
each ticker symbol before interating through the data set. By doing it this way, the analysis would be 
completed much faster than using the nested for loop for earlier.

### Run-time for Each Method and yearValue

Here are the run-times using the original code.

[2017 Original Run-time]<https://github.com/Wall-E28/stock_analysis/blob/master/Resources/VBA_Challenge_2017_Orginial.png> 

[2018 Original Run-time]<https://github.com/Wall-E28/stock_analysis/blob/master/Resources/VBA_Challenge_2018_Original.png>

Here are the run-times using the refactored code.

[2017 Refactored Run-time]<https://github.com/Wall-E28/stock_analysis/blob/master/Resources/VBA_Challenge_2017.png>

[2018 Refactored Run-time]<https://github.com/Wall-E28/stock_analysis/blob/master/Resources/VBA_Challenge_2018.png>

Based on the run-times, it is apparent that the refactored code run about .5 seconds faster than the original code making it more efficient. 

## Summary on Refactoring 

### General thoughts on Refactoring

The major advantage of refactoring code is making the code more efficient. The major disadvantage of refactoring code is that you are taking code that already works and potential making it unusable if you can refactor it correctly. For that reason it is always smart to save your original code just incase you end up not being able to refactor it. 

### Refactoring in VBA Script

The major advantage of refactoring code in VBA script is that you can use as much as of the original code as you want to and can put your new code side by side with your old code using different modules. The major disadvantage of refactoring code in VBA script is that if you do not have a strong understanding of the syntax, you will struggle to refactor your code as the syntax matters so much more when trying to make your code more efficient. 