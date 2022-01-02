# Stock Analysis with VBA

## Overview of Project

### Purpose
The following analysis in intended to provide actionable information regarding several specific stocks within the category of "Green Stocks." The primary analysis evaluates the total daily volume and yearly return for each of the selected stocks in 2017 and 2018.

The stock tickers included in this analysis are: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR.  

A secondary purpose of this report is to provide a macro-enable Excel spreadsheet that can be used for additional analyses in the future. The VBA code is included in the report and discussed below.

## Results

### Analysis of Green Stocks 2017
Based on the assumption that frequently traded stocks will result in accurate stock prices, the total daily volume has been included in this analysis of selected Green Stocks. In the image below there is a table including the total daily volume and yearly return for selected green stocks in 2017.  

IMAGE

In 2017, returns for the selected stocks fell into several categories: negative returns (TERP), modest returns from 0 - 49.9% (AY, CSIQ, HASI, RUN, and SPWR), good returns from 50-99.9% (JKS,VSLR), and significant returns exceeding 100% (DQ, ENPH, FSLR, SEDG). Total daily volume for all stocks exxceeded 100M except DQ and HASI.  

### Analysis of Green Stocks 2018
In the image below there is a table including the total daily volume and yearly return for selected green stocks in 2018.  

IMAGE

In 2018, conversely, all stocks exerpienced negative returns except ENPH and RUN. Four stocks (AY, SEDG, TERP, VSLR) experienced single digit negative returns, while six stocks (CSIQ, DQ, FSLR, HASI, JKS, SPWR) experienced double digit negative returns. ENPH returned 81.9% on its stock price in 2018. RUN returned 84% on its stock price in 2018. Total daily volume was comparable to 2017, with all stocks exceeding 100M except for AY.  

### Green Stocks 2017-2018 Comparison
With the majority of stocks in the Green energy sector experiecing negative returns in 2018, a simple analysis of such returns would steer the buyer away from all but two stocks. Only ENPH and RUN experienced positive returns in both 2017 and 2018, with ENPH being the better performing stock overall. These two stocks warrant further analysis using more advanced metrics to determine whether either should be labeled as a "Buy" at this moment. The other stocks warrant further investigation, as well, but their negative returns in 2018 are cause for concern and required explanation.

### Analysis of VBA Script
The VBA script included in this report can be utilized for this stock analysis only. The script goes through the following processes:  
1. Requests user to choose year for analysis;  
2. Starts a timer to determine length of analysis;  
3. Hard codes the tickers for analysis;  
    IMAGE  
4. Loops through the spreadsheet of the chosen year once per ticker;  
5. Calculates total daily volume and gathers starting and ending price for each ticker;  
    IMAGE  
6. Creates a table displaying the calculated and gathered data for each ticker;  
7. Formats the table based on positive returns (green fill) and negative returns (red fill);  
    IMAGE  
8. Displays a Message Box indicating how long the analysis took.  
    IMAGE  


## Summary
The VBA code is rudimentary and would require further refactoring for use outside of this analysis. The bones of the script, however, are solid.  

### VBA Advantages
The primary advantage to the script as it is currently coded is that the analysis loops through the dataset and gathers the relavent data for each ticker. Additional data could be gathered or calculated with a few lines of code. This advantage is a direct result of the first refactoring of the VBA code.  

### VBA Disadvantages
Currently, the tickers are hard coded into the script. This is the primary detail that would need to be addressed in order to utilize this code for other stocks. A loop gathering unique tickers in each dataset would prove useful. Additionally, it would be useful to present the information gathered in a way that allows comparison by year - that is, to see multiple years side by side. One negative to this refactoring is that the work of refactoring took quite a while and did not result in a great reduction of analysis time, nor did it finish increasing the usability of the code. Further work is necessary before it would provide general usability.  

