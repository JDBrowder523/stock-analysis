# Analysis of Stocks with VBA and Excel

## Overview of the Project
  In this project, we had data from several different stocks over 2017 and 2018.  Through analysis using VBA with Microsoft Excel, we ran a refactored code that outputs the total daily trading volume and the percentage that the stock returned to determine if this refactored code was faster than the original VBA script.
  
## Results and Analysis
### Analysis of Refactored Code
  The code to be refactored, or restructured to improve performance and operation, was copied from a README file given and pasted into a new module in VBA for Microsoft Exel.  The code already had elements that created the sub, created the variables yearValue, and startTime created the headers "Ticker", "Total Daily Volume", and "Return", and also created an array of tickers for each stock ticker in the dataset.(This is an image)---
  The original code also had precoded script that changed the cell color or positive retuens to green and negative returns to red.  This made the final output data much easier to analyze to determine which stocks were high perfomrers and which needed to be avoided.(This is an image)---
  The code that needed to be refactored was layed out with comments describing the code that needed to be written.  Following each step, the following code was produced: (This is an image)---
  
  ### Results
    The original code created using the classwork module had a run time of 0.37 seconds for 2017 stock data analysis and 0.40 seconds for 2018 stock data analysis. The refactored code ran faster and was easier to follow.  The run time for the 2017 stock data analysis with the refactored code was 0.289 Seconds and the run time for the 2018 stock data analysis was also 0.289 seconds.  While this does not seem like much, the refactored code allows for faster programming and improved machine performance.(This is an image) (This is an image)
    
## Summary
### Advantages and Disadvantages of Refactoring Code
  Refactoring any code is helpful because it makes the code easier to read and follow, as well as imprives performance and run time of programs.  It is beneficial to have code that is in its most simple form so other programmers can easily understand what each line is accomplishing. 
  However, refatoring code can be time consuming and confusing.  If you do not have a firm grasp on what each line in the original code is doing, attempting to refactor the code render the program usless or add bugs that decrease performance.
### Advantages and Disadvantages of Refactoring Stock Analysis VBA Script
  By refactoring the stock ananysis VBA code, we were able to increase the performance of the program. Also, the added comments in the refactored code made it easier to understand and will allow for further refactoring for a programmer with more knowledge of VBA. One disadvantage of this refactoring is that it was very time consuming with the outcome being that it produced the same result as the original code.
  
