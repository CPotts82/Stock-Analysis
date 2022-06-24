# Stock-Analysis
## Overview
### Background
  In this analysis I am helping a finance graduate research alternative energy stocks to help his clients invest wisely and diversely.  The clients would like to put all money into one particular stock, DAQO, however research must be done on the DAQO stock as well as other renewable energy stock options to make sure this is a wise choice. The clients have done no research on there own, they just know that they would like to invest in green energy.  
### Purpose
  The purpose of this analysis is to not only research the DAQO stock for the clients but to also research other renewable energy stocks that could possibly be invested in.  This research will consist of finding the Total Daily Volume (how often a stock was traded in a year), as well as the Yearly Return.  The clients are specifically interested in the volume and return for 2018 but we will write code so we can compare 2018 to 2017. I will start with writing code to analyze one stock (DAQO), then build upon that code to analyze 12 different stocks and ultimately refactor the code so that it will work on the entire stock market that will run across the last few years. 
 ## Results of Analysis
 ### Stock Analysis - DAQO and All Stocks
 2017 was a fairly good year for the renewable energy stocks researched.  There was a lot of trade activity with a range starting at 35,796,200 going all the way up to 782,187,000 trades for the year.  The yearly returns were pretty impressive with only 1 of the stocks having a negative return at -7.2%.  The other 11 stocks were all in the positives ranging from 5.5% to 199.4% return.  DAQO had the lowest amount of trade activity but the highest yearly return for the year at 199.4%.
 ![2017](https://user-images.githubusercontent.com/106348899/175434518-72af9cbd-6010-4294-a9a9-9b12f0a2311b.png)
2018 showed a lot of trade activity across the board for all of the stocks researched with the volume ranging from 83,079,900 to 607,473,500 trades for the year.  More trade activity in 2018 was most likely due to the poor returns for the green energy stocks.  There were only 2 stocks that had positive yearly returns for 2018. The ranges for the yearly returns was from -62.6% to 84.0%.  In 2018 DAQO had a little more trade activity but also had the lowest yearly return at -62.6%. 
![2018](https://user-images.githubusercontent.com/106348899/175435050-9ad1ff01-341d-487b-9743-0b5c85abd6b3.png)
### Analysis Procedure/Code
The DAQO analysis was pretty straightforward only analyzing 1 stock but when it came to researching all 12 stocks the coding became a bit more complex.  In order for our code to loop through and pull the Total Volume and Returns per year we had to set up an array with all the tickers for the computer to loop through to pull and compile our data. In the original All Stocks code I ended up using a nested for loop to initialize the ticker, increase volume per ticker and get the starting and ending prices for the return.  


```For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    '5) loop through rows in the data
       Sheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value
           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
           End If
           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
           End If 
       Next j 
```
#### Refactored Code
For the refactored code I used 1 For loop to initialize the ticker Volume and then used another For loop with If/Then statements to begin looping over all the data rows, increase the ticker volume per ticker, find the starting and ending prices as well as increase the ticker index. 


```tickerIndex = i
For i = 0 To 11
    tickerVolume(i) = 0   
Next i   
    ''2b) Loop over all the rows in the spreadsheet.
        rowStart = 2
        rowEnd = Cells(Rows.Count, "A").End(xlUp).Row   
For i = rowStart To rowEnd

        '3a) Increase volume for current ticker       
        If Cells(i, 8).Value = tickerVolume(tickerIndex) Then
               tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(j, 8).Value       
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex 
        If Cells(i, 1).Value = tickerIndex And Cells(i - 1, 1) <> tickerIndex Then
               tickerStartingPrice(12) = Cells(i, 6).Value     
        End If
                
        '3c) check if the current row is the last row with the selected ticker        
        If Cells(i, 1).Value = tickerIndex And Cells(i + 1, 1) <> tickerIndex Then
               tickerEndingPrice(12) = Cells(j, 6).Value
        End If
    
        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickerIndex And Cells(i + 1, 1).Value <> Cells(i - 1, 1).Value Then
            tickerIndex = tickerIndex + 1           
        End If
```
#### Execution times: Original Code vs. Refactored Code
In both analysis of All Stocks a timer was used to see how long it took each script to run.  The original script had times of .2578125 seconds to run the 2017 script and .265625 seconds to run the 2018 script.  The refactored code took much less time for both years. I did notice that in order to get the shorter run time I had to run the code from VBA.  When the analysis are run from the button on the worksheet the times are similar to the original scripts.  

![VBA_Challenge_2017](https://user-images.githubusercontent.com/106348899/175437784-fd5580ec-8ff8-4875-a20d-299ffb29e6f0.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/106348899/175437800-41f5bdbe-452b-491c-a891-3a6d85dd20cb.png)
      
 ## Summary
 ### Advantages and Disadvantages of Refactoring Code
 I believe that refactoring code can be very advantageous as you are cleaning up, simplifying and improving upon code that has already been written.  Writing code is like writing a paper, there are multiple drafts before you get your final copy.  Code can always be improved upon. The disadvantage to refactoring code is, especially if it is someone else's code, trying to figure out anything that is not clear.  Trying to debug where you do not have background knowledge.  It can be challenging but overall I believe it has the potential to be mostly advantageous. 
### Pros/Cons

