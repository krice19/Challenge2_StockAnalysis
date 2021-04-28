# Challenge2_StockAnalysis
Challenge 2 for Stock Analysis - VBA
## Overview
The purpose of this analysis is to refactor a solution code to make it more efficient and run faster, while obtaining valuble insights about stock information.  I am looking to refactor this code so it can be used on larger data sets.  The first attempt at the code was to gather information on twelve different stocks from specific years through data spreadsheets.  From the data, I pulled information on the stock volume and return, depending on what year the user would input.  In this code, I used a nested for loop, which looped through each row of data for each of the twelve stocks.  The times it took for execute the code for data sheets 2017 and 2018 before the refactor are below: <br />
<img width="1440" alt="VBA_Challenge_2017_Before" src="https://user-images.githubusercontent.com/63257696/116460442-12f96780-a835-11eb-8cc3-6f9181b747c8.png">
<img width="1440" alt="VBA_Challenge_2018_Before" src="https://user-images.githubusercontent.com/63257696/116460461-168cee80-a835-11eb-921b-dadba991213d.png">
<br />
You can see that the code took about .8 and .7 seconds to execute.  While this code worked well for a data sheet of 12 stocks, it may not work as well for thousands of stocks.  In this analysis,  I will refactor my first script to only loop through the rows once.  I will evalute the results of my code and determine how the refactor helped it execute more efficiently.  The output of the code will then help me compare stock information, which can provide valuable insights into investments.
## Results
In my refactor,  I looped through each row of data once in order to speed up the code execution time.  I did this be creating a variable called "ticker index", and set it to 0.  The ticker index will be used to loop through the arrays made for ticker name, ticker volume, ticker starting price, and ticker ending prices.  I created a for loop to initialize stock volume item in the array to 0.  I ended that for loop and created another for loop to loop through each row in the data sheet, and assign cell values in the array based off certain conditions using "If then" statements.  The last "if then" statement in the for loop increased the ticker index by 1, indicating to assign values to the next ticker name.  I ended that for loop, and created one last for loop to output the array values on a spreadsheet.  An example for my code is below:<br />
    
    '1a) Create a ticker Index and set to 0
    
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickersVolume(12) As Long
    Dim startingPrices(12) As Single
    Dim endingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'the loop will increase the ticker index and set all volumes in the array to 0
    
    For i = 0 To 11
        tickersVolume(tickerIndex + i) = 0
    Next i

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
   
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickersVolume(tickerIndex) = tickersVolume(tickerIndex) + Cells(i, 8).Value
        End If
       
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'if it is the first row, set the value to the starting price
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            startingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'if it is the last row, set the value to the ending price
        'If  Then
            
                 If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                endingPrices(tickerIndex) = Cells(i, 6).Value
                End If

            '3d Increase the tickerIndex
            'the ticker index will increase by 1 for the same criteria as the ending price, so it it is the last row of the current ticker
            
            
               If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                 tickerIndex = tickerIndex + 1
                End If
            
        'End If
    
    Next i
 
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    'the for loop will have i loop through the arrays andplace the values in the cells stated below
    'the row of the cell we are placing the value in will also increase by i increase
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickersVolume(i)
        Cells(4 + i, 3).Value = endingPrices(i) / startingPrices(i) - 1
        
After my refactor with the above code, I was able to execute the analysis in a shorter amount of time, as seen below: <br />
<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/63257696/116462008-fbbb7980-a836-11eb-9ef4-76087aa97f19.png">
<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/63257696/116462020-feb66a00-a836-11eb-96b7-2e9a124e5a46.png">
The execution time sped up by .6-.7 seconds, which will help when working with larger data sets.<br />
My code resulted in the following output on the spreadsheet: <br />
<img width="370" alt="VBA_Challenge_2017_Results" src="https://user-images.githubusercontent.com/63257696/116462620-a6cc3300-a837-11eb-98f3-317d3f3de1bd.png">
<img width="390" alt="VBA_Challenge_2018_Results" src="https://user-images.githubusercontent.com/63257696/116462637-ab90e700-a837-11eb-9af3-a473bdfd2027.png">
2017 proved to be a better year for stock return than 2018, as most of the returns are negative in 2018.  The two stocks that remained with a positive return in both years are RUN and ENPH.  The rest of the stocks dropped or remained with a negative return.  The RUN stock was the only stock with a return that increased from 2017 to 2018.  In 2017, it resulted in a 5% return, while in 2018, the return jumped to 84%.  The ENPH stock return remained positive, but fell from about a 100% return to 84%.
## Summary
An advantage of refactoring code is to make it run more efficiently and faster so that it can be used on larger data sets.  The more data on a spreadsheet, the longer VBA has to loop through all the rows, therefore slowing down the process.  This applies to my stock analysis because the refactor ran a lot faster than the original script.  If I included more data, the refactored script would be more efficient.  Disadvantgaes of refactoring is that a new way to run the code may not always be the easiest think of and it may not apply if a data set is set out in a specific way.  This applies to my stock analysis because to first way to think of looping through rows is through a nested for loop, and the idea of putting the data into arrays was not my first thought when writing the script.  Additionally, the refactor of the code would not work if the stock information was in a random order.  For example, the refactored code loops through each row once.  When it meets conditions for the volume, starting praice, and ending price, it stores those values in the array.  Then the ticker index increases if its the last ticker name on the sheet. This refactor works because the ticker data is grouped in order, so the code will recognize when the ticker name changes.  If the tickers were not grouped in order, then the correct information would not be assigned in the array.

