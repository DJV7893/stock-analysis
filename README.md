# stock-analysis
---
## Overview of Project
---
### Purpose 

Steve, a recent graduate with a degree in finance, asked us to analyze a list of green energy stock data contained in an excel file to diversify his client’s portfolio. Using the excel extension Visual Basic Applications (VBA), we wrote code to automate functions such as: analyses determining total yearly volume and return for each stock, formatting output tables to make them readable and highlight differences in stock performance, creating user friendly buttons to run these analyses, and measure how fast our code will compile these results. Now, Steve wants to expand his dataset to include the entire stock market over the last few years. In order to do so I will be refactoring our solution code to loop through all the data one time to collect the same information but taking fewer steps and decreasing script run time. I will also look at the current data collected to assess stock performance.

---
## Results
---
### Stock Performance

In the 2017 and 2018 stock data that was output into two separate tables we can see a significant difference in yearly stock performance. In 2017 all stocks except for TERP had a positive return especially DQ, ENPH, FSLR and SEDG all of which had a return of more than 100%. However, in 2018 we can see a steep decline in stock performance with all but 2 of the stocks, ENPH and RUN, having a negative yearly return.

<img width="486" alt="2017_vs_2018_StockPerformance" src="https://user-images.githubusercontent.com/99817571/156820400-87ccb2e1-412f-4929-9774-3c3a18528118.png">

### Code and Execution Time
The main difference between the refactored and original code is the creation of arrays for the calculation of Total Volume, Starting Price and Ending Price for each stock. These arrays allow us to eliminate the nested loop and capture in memory the stock information. As we can see in the altered script, the use of an index allows us to loop through the different arrays for the calculation of the values. In the original script we only used an array for the stock ticker and all the calculations are done through the inner loop and before going to the next stock index the values are outputted into the table. Meanwhile, in the refactored code once the values in all 3 of the output arrays are calculated, we then have to create a loop to display the stock ticker, total volume and yearly return in the table.

    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
    tickerVolumes(i) = 0
    
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
            End If
            
            
        '3c) check if the current row is the last row with the selected ticker
         
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        
            '3d Increase the tickerIndex.
        
        tickerIndex = tickerIndex + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

The changes in the script resulted in a significant decrease in execution time for both output tables. For both the 2017 and 2018 stock data the execution time of the code was reduced from .52 seconds to .11 seconds.

<img width="536" alt="Run_Time_2017s" src="https://user-images.githubusercontent.com/99817571/156822048-c46e6226-2c8f-40f6-9b14-7efec48fa1f4.png">

<img width="537" alt="Run_Time_2018s" src="https://user-images.githubusercontent.com/99817571/156822094-1210e16a-6ec5-4f61-9011-478d0ce09e55.png">

---
## Summary
---
Refactoring allows us to go through our original code in order to scrutinize elements of our script and identify if there are more efficient ways of collecting and outputting the same information. If we stay put with our original code, we could run into issues in the future if more data or different variables are incorporated into the dataset. This could potentially slow down our ability to produce the desired data. However, in the process of refactoring our code we can potentially distort or halt entirely the output of information if our syntax is not accurate. It can be a tedious process to refactor code if one is not attentive throughout when removing or improving the logic of certain steps throughout the script if they do not flow with the rest of the script.

As we saw when refactoring the Stock Analysis code, improving the logic by removing nested loops for multiple arrays reduced the execution time by .41 seconds. But the process of editing that portion of the script was tedious as I ran into several errors when updating the code for it to sequentially store the data into each array before proceeding to the next index. Had I not ensured that the updated portions of the script worked seamlessly with the rest of the Stock Analysis code, then I wouldn’t be able to even produce the original output tables. With a careful attention to syntactical changes the new script gave the desired result of improved execution time without output errors.
