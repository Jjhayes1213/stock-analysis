# Stocks Analysis

## Overview of Project
The client had originally chosen a stock ticker "DQ" for potential investments.  After analyzing the "DQ" stocks return value for 2017 and 2018, it was determined that this may not be the best investment choice.  The search for investment choices was expanded to tweleve stock tickers, including the original stock reviewed "DQ".  All 12 stocks have been analyzed for both the total daily volume and the return percentage for the years 2017 and 2018.

### Results
After reviewing the data for the 12 stocks, it appears that all but 1 stock "TERP" had performed well in 2017.  Including the original inquiry for "DQ" which showed 199.4% return.  On review of the 2017 data tickers "DQ", "ENPH", "FSLR", and "SEDG" were the highest performers for the year.  However we were given the data for the stock performance on these 12 stocks for both 2017 and 2018.  VBA code was written to include the ability to run the data for both years based on the users input.

yearValue = InputBox("What year would you like to run the analysis on?")
Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks(" + yearValue + ")"

Also very important was for the code to loop through the tickers without too much repetition. An array was created so the code code look for the data associated with the Ticker numbers listed.  And a tickerIndex was created to be input for those arrays into the if then statements.

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

 tickerIndex = 0   

 If Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

These techiniques helped for the code to run quickly, within hundreths of a second.  A previous version of this code was written labeled "All Stocks Analysis".  Though the code was sutibale, the return of values to the chart after user input was within tenths of a second.

Now as mentioned earlier, the code was written to take user input on which year to run.  When the user runs the report for the year 2018, it appear that all of the stocks performed poorly except for "ENPH" and "RUN".  Both of these stocks had over 80% return in 2018.

By reviewing that "ENPH" had a high rate of return in both 2017 and 2018, I would recommend this stock option for the client.


### Summary
Having performed this code originally and then having wrote the script again by refactoring the code, I do see some advantages and disadvantages.  Though I found the original code less difficult to write, as I look back it is harder to follow and would be more difficult to add for loops and if then statements into the work without rewritting the code.  

I found the flow of the refractored code to be more organized.  The headers make it very easy to reference back and read through the sequencing.  But it is also easier to see how you can make minor changes within the script with out having to re-write major pieces of the code.
