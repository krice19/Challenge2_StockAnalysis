# Challenge2_StockAnalysis
Challenge 2 for Stock Analysis - VBA
# Overview
The purpose of this analysis is to refactor a solution code to make it morre efficient and run faster, while obtaining valuble insights about stock information.  I am looking to refactor this code so it can be used on larger data sets.  The first attempt at the code was to gather information on twelve different stocks from specific years through data sheets.  From the data, I pulled information on the stock volume and return, depending on what year the user would input.  In this code, I used a nested for loop, which looped through each row of data for each of the twelve stocks.  The times it took for execute the code for data sheets 2017 and 2018 before the refactor are below: <br />
<img width="1440" alt="VBA_Challenge_2017_Before" src="https://user-images.githubusercontent.com/63257696/116460442-12f96780-a835-11eb-8cc3-6f9181b747c8.png">
<img width="1440" alt="VBA_Challenge_2018_Before" src="https://user-images.githubusercontent.com/63257696/116460461-168cee80-a835-11eb-921b-dadba991213d.png">
<br />
You can see that the code took about .8 and .7 seconds to execute.  While this code worked well for a data sheet of 12 stocks, it may not work as well for thousands of stocks.  In this analysis,  I will refactor my previous code to only loop through the rows once.  I will evalute the results of my code and determine how the refactor helped it execute more efficiently.  The output of the code will help me compare stock information, which can provide valuable insights into investment.
# Results
In my refactor,  I looped through each row of data once in order to speed up the code execution time.  I did this be creating a variable called "ticker index", and set it to 0.  The ticker index will be used to loop through the arrays made for ticker index name, ticker volume, ticker starting price, and ticker ending prices.  I created a for loop to initialize stock volume item in the array to 0.  I ended that for loop and created another for loop to loop through each row in the data sheet, and assign cell values to values in the array based off certain conditions using an "If then" statement.  The last "if then" statement in the for loop increased the ticker index by 1, indicating to assign values to the next ticker.  I ended that for loop, and created one last for loop to output the array values on a spreadsheet.  After my refactor, I was able to execute the code in a shorter amount of time, as seen below: <br />
<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/63257696/116462008-fbbb7980-a836-11eb-9ef4-76087aa97f19.png">
<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/63257696/116462020-feb66a00-a836-11eb-96b7-2e9a124e5a46.png">
The execution time sped up by .6-.7 seconds, which will help when working with larger data sets.<br />
My code resulted in the following output on the spreadsheet: <br />


# Summary


