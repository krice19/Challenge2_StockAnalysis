# Challenge2_StockAnalysis
Challenge 2 for Stock Analysis - VBA
## Overview
The purpose of this analysis is to refactor a solution code to make it morre efficient and run faster, while obtaining valuble insights about stock information.  I am looking to refactor this code so it can be used on larger data sets.  The first attempt at the code was to gather information on twelve different stocks from specific years through data sheets.  From the data, I pulled information on the stock volume and return, depending on what year the user would input.  In this code, I used a nested for loop, which looped through each row of data for each of the twelve stocks.  The times it took for execute the code for data sheets 2017 and 2018 before the refactor are below: <br />
<img width="1440" alt="VBA_Challenge_2017_Before" src="https://user-images.githubusercontent.com/63257696/116460442-12f96780-a835-11eb-8cc3-6f9181b747c8.png">
<img width="1440" alt="VBA_Challenge_2018_Before" src="https://user-images.githubusercontent.com/63257696/116460461-168cee80-a835-11eb-921b-dadba991213d.png">
<br />
You can see that the code took about .8 and .7 seconds to execute.  While this code worked well for a data sheet of 12 stocks, it may not work as well for thousands of stocks.  In this analysis,  I will refactor my previous code to only loop through the rows once.  I will evalute the results of my code and determine how the refactor helped it execute more efficiently.  The output of the code will help me compare stock information, which can provide valuable insights into investment.
## Results
In my refactor,  I looped through each row of data once in order to speed up the code execution time.  I did this be creating a variable called "ticker index", and set it to 0.  The ticker index will be used to loop through the arrays made for ticker name, ticker volume, ticker starting price, and ticker ending prices.  I created a for loop to initialize stock volume item in the array to 0.  I ended that for loop and created another for loop to loop through each row in the data sheet, and assign cell values to values in the array based off certain conditions using an "If then" statement.  The last "if then" statement in the for loop increased the ticker index by 1, indicating to assign values to the next ticker.  I ended that for loop, and created one last for loop to output the array values on a spreadsheet.  An example for my code is below:<br />
 
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        Cells(i, 1).Value = tickerName
        Cells(i, 8).Value = currentVolume
       
        '3a) Increase volume for current ticker
             If tickerName = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + currentVolume
            End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
            If tickerName = tickers(tickerIndex) And tickerName - 1 <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
             If tickerName = tickers(tickerIndex) And tickerName + 1 <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            End If
            
 
            '3d Increase the tickerIndex.
             If tickerName = tickers(tickerIndex) And tickerName + 1 <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
    Next i
 
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stock Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
     
    Next i



After my refactor with the above code, I was able to execute the analysis in a shorter amount of time, as seen below: <br />
<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/63257696/116462008-fbbb7980-a836-11eb-9ef4-76087aa97f19.png">
<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/63257696/116462020-feb66a00-a836-11eb-96b7-2e9a124e5a46.png">
The execution time sped up by .6-.7 seconds, which will help when working with larger data sets.<br />
My code resulted in the following output on the spreadsheet: <br />
<img width="370" alt="VBA_Challenge_2017_Results" src="https://user-images.githubusercontent.com/63257696/116462620-a6cc3300-a837-11eb-98f3-317d3f3de1bd.png">
<img width="390" alt="VBA_Challenge_2018_Results" src="https://user-images.githubusercontent.com/63257696/116462637-ab90e700-a837-11eb-9af3-a473bdfd2027.png">
2017 proved to be a better year for stock return than 2018, as most of the returns are neegative in that year.  The two stocks that remained with a positive return are RUN and ENPH.  The rest of the stocks dropped or remained with a negative return.  The RUN stock was the only stock with a return that increased from 2017 to 2018.  In 2017, it resulted in a 5% return, while in 2018, the return jumped to 84%.  The ENPH stock return remained positive, but fell from about a 100% return to 84%.
## Summary
An advantage of refactoring code is to make run more efficiently and faster so that it can be used on larger sets.  The more data on a spreadsheet, the longer VBA has to loop through all the rows, therefore slowing down the process.  This applies to my stock analysis because the refactor ran a lot faster than the original script.  If I included more data, the refactored script would be more efficient.  A disadvantgaes of refactoring is that a new way to run the code may not always be the easiest think of and it may not apply if a data set is laid out in a specific way.  This applies to my stock analysis because to first way to think of looping through rows is through a nested for looping, and the idea of putting the data into arrays was not my first thought when writing the script.  Additionally, the refactor of the code would not work if the stock information was in a random order.  For example, the refactored code loops throigh each row once.  When it meets conditions for the volume, starting praice, and ending price,grouped init sotres those values in the array.  Then the ticker index increases if its the last ticker name on thesheet. This refactor works because the ticker data is grouped in order. Put if the tickers were not grouped in order, then the correct information would not be assigned in the array.

