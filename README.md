# Stock Analysis with VBA

## Overview of Project

### Purpose
The following analysis in intended to provide actionable information regarding several specific stocks within the category of "Green Stocks." The primary analysis evaluates the total daily volume and yearly return for each of the selected stocks in 2017 and 2018.

The stock tickers included in this analysis are: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR.  

A secondary purpose of this report is to provide a macro-enable Excel spreadsheet that can be used for additional analyses in the future. The VBA code is included in the report and discussed below.

## Results

### Analysis of Green Stocks 2017
Based on the assumption that frequently traded stocks will result in accurate stock prices, the total daily volume has been included in this analysis of selected Green Stocks. In the image below there is a table including the total daily volume and yearly return for selected green stocks in 2017.  

![Imgur](https://i.imgur.com/EhZOBvn.png)

In 2017, returns for the selected stocks fell into several categories: negative returns (TERP), modest returns from 0 - 49.9% (AY, CSIQ, HASI, RUN, and SPWR), good returns from 50-99.9% (JKS,VSLR), and significant returns exceeding 100% (DQ, ENPH, FSLR, SEDG). Total daily volume for all stocks exxceeded 100M except DQ and HASI.  

### Analysis of Green Stocks 2018
In the image below there is a table including the total daily volume and yearly return for selected green stocks in 2018.  

![Imgur](https://i.imgur.com/ZJ92bZD.png)

In 2018, conversely, all stocks exerpienced negative returns except ENPH and RUN. Four stocks (AY, SEDG, TERP, VSLR) experienced single digit negative returns, while six stocks (CSIQ, DQ, FSLR, HASI, JKS, SPWR) experienced double digit negative returns. ENPH returned 81.9% on its stock price in 2018. RUN returned 84% on its stock price in 2018. Total daily volume was comparable to 2017, with all stocks exceeding 100M except for AY.  

### Green Stocks 2017-2018 Comparison
With the majority of stocks in the Green energy sector experiecing negative returns in 2018, a simple analysis of such returns would steer the buyer away from all but two stocks. Only ENPH and RUN experienced positive returns in both 2017 and 2018, with ENPH being the better performing stock overall. These two stocks warrant further analysis using more advanced metrics to determine whether either should be labeled as a "Buy" at this moment. The other stocks warrant further investigation, as well, but their negative returns in 2018 are cause for concern and required explanation.

### Analysis of VBA Script
The VBA script included in this report can be utilized for this stock analysis only. The script goes through the following processes:  
1. Requests user to choose year for analysis;  
2. Starts a timer to determine length of analysis;  
3. Hard codes the tickers for analysis;  
    ```
    Sub AllStocksAnalysisRefactored()
        Dim startTime As Single
        Dim endTime  As Single

        yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer
    
        'Format the output sheet on All Stocks Analysis worksheet
        Worksheets("AllStocksAnalysis").Activate
    
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

        'Initialize array of all tickers
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
        ...
    ```  
    
4. Loops through the spreadsheet of the chosen year once per ticker;  
5. Calculates total daily volume and gathers starting and ending price for each ticker;  
    ```
    ...
        'Activate data worksheet
        Worksheets(yearValue).Activate
    
        'Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
        '1a) Create a ticker Index
        tickerIndex = 0

        '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
        '2a) Create a for loop to initialize the tickerVolumes to zero.
        For zz = 0 To 11
            tickerVolumes(zz) = 0
        Next zz
                
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
            '3a) Increase volume for current ticker
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            End If

        Next i
    ...
    ```  
    
6. Creates a table displaying the calculated and gathered data for each ticker;  
7. Formats the table based on positive returns (green fill) and negative returns (red fill);  
    ```
    ...
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("AllStocksAnalysis").Activate
    
    For i = 0 To 11    
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
    
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd        
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen            
        Else
            Cells(i, 3).Interior.Color = vbRed            
        End If        
    Next i
    ...
    ```  

8. Displays a Message Box indicating how long the analysis took.  
    ```
    ...
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
    ```  
    
## Summary
The VBA code is rudimentary and would require further refactoring for use outside of this analysis. The bones of the script, however, are solid.  

### VBA Advantages
The primary advantage to the script as it is currently coded is that the analysis loops through the dataset and gathers the relavent data for each ticker. Additional data could be gathered or calculated with a few lines of code. This advantage is a direct result of the first refactoring of the VBA code. Additionally, the time for running the macro was reduced approximately 5x. This is not significant in this small dataset, but with a larger dataset, the time saved could be significant.  

### VBA Disadvantages
Currently, the tickers are hard coded into the script. This is the primary detail that would need to be addressed in order to utilize this code for other stocks. A loop gathering unique tickers in each dataset would prove useful. Additionally, it would be useful to present the information gathered in a way that allows comparison by year - that is, to see multiple years side by side. One negative to this refactoring is that the work of refactoring took quite a while and did not result in a great reduction of analysis time (for this dataset), nor did it increase the usability of the code beyond this dataset. Further work is necessary before it would provide general usability.  

The first version of this VBA code took 0.56 seconds for 2017 and 0.57 seconds for 2018.
![Imgur](https://i.imgur.com/jnQ1smx.png?1) ![Imgur](https://i.imgur.com/hTbi6bC.png?1)  

The refactored code took 0.51 seconds for 2017 and 0.50 for 2018.
![Imgur](https://i.imgur.com/2SfIkww.png) ![Imgur](https://i.imgur.com/II6XteJ.png)
