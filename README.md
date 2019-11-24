# Original Stock Analysis
## Purpose of the Original Analysis
Using VBA codes to automate and perform basic analysis on stocks in the datasheet for Steve.  

**Analysis elements:**
Basic analysis included summarizing each stock’s annual total volume, yearly return percentage, starting price at the beginning of the year, and ending price at the end of the year.  Outputs are formatted for better presentation, and buttons added to allow Steve to run analysis easily without using the Developer function.

# Challenge
## Purpose of the Challenge
Refactoring the original codes to streamline and improve the logical flow.  This allows larger dataset to be analyzed faster in the future.

## Key Refactoring Implemented in This Challenge
1.	**Improve processing efficiency:**  Original codes require each stock ticker to loop through the entire data rows.  For example, that would be 12 tickers X 3012 rows X 3 conditions for 2018 data. The refactored codes only go through the entire datasheet once.

2.	**Switching order of loops:**  By changing the order of loops, the program now goes down the datasheet only once.  As the program goes through each row, it runs through the same 3 conditions as before:  adding current daily volume to total volume if it’s the same ticker, setting starting price if current row has a different ticker than previous row, and setting ending price if current row has a different ticker than the next row.  By switching the loop order, the processing is now 3012 rows X 3 conditions for 2018 data, which is much more efficient.

3.	**Remove the nesting loops:**  As the program reaches the condition that the next row has a different ticker, it will create an output row for the current ticker first, then resets the Total Volume to zero and increases the tickerIndex by one for the next row.  The original codes used two iterators to loop through each ticker and the datasheet.  The refactored codes only use one iterator in the For Loop.  

4.	**Change in output elements:**  Steve did not like the Return variable in the original codes.  The refactored codes removed the Return variable, and showed the Starting and Ending Prices instead.
