# VBA Challenge

## Overview of Project
This analysis was to compare stock price perfomance using year data from 2017 and 2018 and automate the analysis using VBA. The stock compared green energy stock and compared their returns to see what stock would be  a good investments. The sheet would also highlight which stocks were positive returns and which stock had negative returns from the yearly data.

### Purpose
The analysis is for Steve's parents because they wanted to invest in green energy stocks and wanted to see what would be the best investment for them. After the initial code I refactored it to run more efficiently and to use less memory. 

## Results
The first iteration of the VBA code ran through a set of loops and goes through each ticker symbol and assigned a value to index it from the array and reset the volume each time the loop runs. 

   For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0

In the refactored code instead of assinging tickers and looping through the variable I set the loop up to reset the tickerVolume and to go through the loop with one ticker and at the end of the loop adds one to the index value and then resets the code.

For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
    
    '2b) loop over all the rows
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                    
                    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
            End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                
                    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
         
         
         '3d) ticker index increasing by 1
         
                    tickerIndex = tickerIndex + 1
            
            End If

        Next j

This section of the code made the code run time significantly faster comparing it to the first code. The first code ran at around .73 seconds and after the refactor it ran at around .101 seconds.


<img = src"Resources/green_stock_code_time.png">

<img = src"Resources/VBA_Challenge_2017.png">

## Summary
 1). The advantages of refactoring code was to improve the efficiency and allow to code to run quicker by improving the flow of the code and the functions used. A disadvantage of refactoring is that in trying to improve the code at certain points I ran into debugging errors. 

 2). The Pro was that I was able to cut the time my code took to ran significantly after I refactored it. The con that apply to my VBA code was in the beggining when I was trying to improve the loop function I ran into overflow and compiling errors. After adjustments I was able to refactor the code in order to improve the time. 

