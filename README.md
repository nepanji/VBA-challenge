# VBA-Challenge
## Purpose
By completing this homework assignment, I was able to demonstrate my ability to use VBA scripting to analyze real stock market data in Microsoft Excel.

## Background
At the start of this assignment, I was provided with 3 years of stock data in Microsoft Excel.  Each year of stock data was loaded in separate worksheets of the Excel Workbook file. To accomplish an analysis of the stock market data, I created a VBA script to automate all of the tedious tasks required to conclude the analysis.
The completed assignment includes:
•	3 Screenshots of the results
•	Separate VBA script files
•	README file

## Scripting Environment
I completed this assignment in the Windows 10, Microsoft 365, Version 16.51 environment. I ran the script on macOS Big Sur 11.2.2.

## Scripting Summary
The Excel file was not saved to this repository because it is very large and takes up a ton of space. Therefore, I created a script file called StockDataAnalysisCode.bas located in the VBA_Challenge folder of this repository. 
Per the assignment, the script runs as requested. 
•	The script will loop through all the stocks for one year and output the following information:
    o	The ticker symbol.
    o	Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    o	The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    o	The total stock volume of the stock.
    o	Conditional formatting will highlight positive change in green and negative change in red.
•	The script will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
•	The script executes on each worksheet and needs to be run only once

## Result Illustrations

### Stock Data (2014)

### Stock Data (2015)

### Stock Data (2016)