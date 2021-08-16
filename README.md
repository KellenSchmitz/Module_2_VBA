# Module_2_VBA

## Project Overview: Refactoring VBA Script for Stock Market Analysis
Phase 1 of the project was to analyze the stock returns of a basket of green stocks in 2017 and 2018. Daily returns and volume were provided for 12 green stocks for the years 2017 and 2018. Code was developed to conduct this analysis and provide a summary on each stock showing the total volume, and price return for each year. 

In Phase 2 of the project, that code was refactored to "improve efficiency by taking fewer steps, using less memory, and improving the logic of the code to make it easier for future users to read." To determine if the refactoring process was successful, code run-time was measured and compared to the before and after versions of the code.

## Project Results
Refactoring led to measurable improvements in the run-time of the stock analysis code. 

Run Times Before Refactoring

<img width="260" alt="before refactor_2017" src="https://user-images.githubusercontent.com/86166117/129621491-4184562a-8698-46b4-9169-6c1b6688d531.png">
<img width="263" alt="before refactor_2018" src="https://user-images.githubusercontent.com/86166117/129622312-d4fc3316-fc07-4fe6-ad8d-eaf8cd4352ec.png">



Refactored Run Times

<img width="280" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/86166117/129621006-aea23e13-db20-4d2a-ae25-cfa8fb37c864.png">
<img width="281" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/86166117/129621054-5d71a257-09a0-4064-9012-c36c43a18ede.png">


One of the refactoring improvements included adding an index to call each ticker. This enhanved code flexibility and reduced resource usage.

'''

    'Ticker index
    tickerIndex = 0

    'Initialize variables for volume, starting price and ending price
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single

    'Loop through the tickers and initialize volume to zero
    For i = 0 To 11
    tickerVolume(i) = 0
    
    Next i

    For i = 2 To RowCount
        'increase volume ofcurrent ticker to find year total volume
        tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
        
       'check for starting price for current ticker
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        
        'check for ending price of current ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next i
'''

## Summary
After refactoring, the code runs in 1/3 the time and is now more flexible. 

Refactoring the code made it easier to read and improved resource usage, but there is the risk that bugs could be introduced during this process. It is also worth noting that the incremental revisions to the original code base may not be productive in the long run, if the original code has fundamental limitations.

Please see the 
Module_2_VBA Refactored.xlsm file for the refactored code.
