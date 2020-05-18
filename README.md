# VBA-challenge
VBA-challenge stock market data

The files uploaded
1. wall_streat_challenge.vbs (macro to run - VBAProject.Module1.Stock_Analysis)
2. Screenshots of results corresponding to 2016, 2015 and 2014
3. Excel sheet with result

Note:- 
Logic is modularised as functions and subroutines

Main Subroutine
Stock_Analysis - In Module1

Supporting Subroutines and functions called within the main subroutine
Functions
 GetAgg
 InArray
 WhereInArray
Subroutine
 Formatting
 

Order of excetion - O(No of rows in the sheet) * No of Sheets
Time taken on the test file over all the sheets - < 2 minutes
Time taken on the main file over all the sheets - About 30 minutes

Extra feeture:
A section would be added if any of the ticker has an opening value = 0
To handle problems caused by division by zero, %change of those tickers will be dispalyed as 0%
However, such a case didn't happen yet in testing
