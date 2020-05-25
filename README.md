# The VBA of Wall Street
Use of  VBA scripting to analyze real stock market data. 

<div style="text-align:center"><img src="images/stockmarket.jpg"></div>

- Create a script that will loop through all the stocks for one year and output the following summary information.

  - The ticker symbol.

  - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  - The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  - The total stock volume of the stock.

- Do conditional formatting that will highlight positive change in green and negative change in red.

- The olution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:


<div style="text-align:center"><img src="images/hard_solution.png"></div>

- Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.


#Solution 1:
VBA code - wall_streat_challenge_solution1.vbs
Run the subroutine Stock_Analysis

<div style="text-align:center"><img src="images/Solution1.png"></div>


- **This solution does not use any workbook application function**
- **It loops through the entire data once, get the summary and loops through the small summary once to get the toppers and floppers.** 
- **This solution works even if the data is not sorted.**
- It loops through all the sheets and gives the result.
- **One extra feature is added to keep track of division by zero error
  - If any ticker has Opening Value = 0, Percent Change becomes Infinity and results in Division by Zero Error
  - Under the occurance of division by zero, Percent Change is taken as 0% and corresponding tickers are logged in the sheet  and highlighed
  - **One Ticker (PNTL) has this problem - happens in 2015 and 2014 years
- **Code returns the total run time taken


As this solution uses arrays that needs redimensioning, it takes about an hour to run. Hence, I don't recommend this one

#Solution 2:
VBA code - wall_streat_challenge_solution2.vbs






