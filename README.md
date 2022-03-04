#  **Analyzing Green Energy Stocks**


## *Overview of Project*


#### A friend and client named Steve has requested assistance in analyzing stock data for his clients. His clients are interested in investing in green energy and have decided to invest their money in the DQ stock. Before advising further, Steve required a more thorough analysis to be completed to ensure the safety of his clients’ investment. By using VBA, a script was created to analyze stock data from whichever year Steve inputs, and then further refactored to make that code more efficient. In doing so, Steve can now make informed decisions about which green energy stock to recommend to his clients.


## *Results*


### Code Breakdown
When writing the code for Steve to use to analyze his stocks, there were 3 main parts of the code that allowed the script to glean various significant data points from the data set.

The first important line of code was establishing the total volume of the stock. This indicates the number of shares of the strength, which some argue correlate to the strength of the stock. The line of code that corresponds to this calculation is as follows:
```
If Cells(j, 1).Value = tickerIndex Then
tickerVolumes = tickerVolumes + Cells(j, 8).Value
End If
```

Additionally, the code needed to find the line of data that corresponded to the beginning of the year in order to calculate the starting price of the stock. To do this, the script is made to look at the line of data above where the stock is listed to see if line is different from the stock being calculated. The starting price of the stock is later compared to the ending price to determine the success of individual stocks. The code needed to perform this function can be seen below:
```
If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
tickerStartingPrices = Cells(j, 6).Value
End If
```

Similarly, the ending price of the stock is determined by looking at the line of data below the stock to see if the value is different. If the stock of the following row is in fact different, this signals to the code that this is the last line of data for a particular stock. An example of this code is shown underneath:
```
If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
tickerEndingPrices = Cells(j, 6).Value
End If
```

### Stock Recommendations
Based on data from the years 2017 and 2018, Steve’s clients should consider investing in either ENPH or RUN, as they each showed growth in both 2017 and 2018. More precisely, ENPH showed a return of 129.5% in 2017 and 81.9% in 2018, whereas RUN produced a return of 5.5% in 2017 and 84.0% in 2018. The analyses for both years can be seen below, along with the time it took for the code to complete the analysis.


![VBA_Challenge_2017](https://user-images.githubusercontent.com/99554642/156848683-f9711fef-40e5-4668-b38f-abc8e0d6127b.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/99554642/156848688-17ea85bf-3ef8-4cf5-ac43-d0a83c6e2f58.png)


## *Summary*
Refactoring code comes with both advantages and disadvantages. The most obvious benefit to refactoring code is to increase efficiency. Refactored code can run faster by analyzing data in a more concise manner. Refactored code can also take up less storage, which can be very valuable. Conversely, refactored code can be a time-consuming process as it can be difficult to figure how exactly to make the code perform better. It might also affect further use of the code as features present in the unfactored code might be needed in the future.

In creating a script for Steve, the main advantage of refactoring was to decrease the time it takes to analyze the data. This is important for Steve as he needs his analyses to be prompt and run concisely. Also, the refactored script runs without looping through unneeded data points. The main disadvantage experienced in refactoring this code was the time it took to make this code run correctly. While the original script produced the same results, it took several attempts to create a working refactored code. 
