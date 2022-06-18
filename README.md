# Module 2: Stocks Analysis

### Project Overview
The project aims to help Steve analyze the performance of multiple green energy stocks to determine the best investment options to recommend to his parents. To perform this analysis, we used the flexible VBA code to run for multiple stocks simultaneously to collect certain stock information in the years 2017 and 2018. My main focus was to refactor the original VBA script to improve its efficiency and explain my findings.
### Results
The tables below contain the total daily volume and the annual rate of return for each of the given 12 green energy stocks:

![Stock_Performance_2017](https://user-images.githubusercontent.com/100629325/174405334-c58d302d-6b76-430b-99c5-6fcd1fac6de7.png)![Stock_Performance_2018](https://user-images.githubusercontent.com/100629325/174405351-3c9b2188-4fd8-4c45-8799-541e9ddab88e.png)
- **The daily stock volume** is the total number of shares traded within a day and measures how actively stock is traded.  Generally, you want to look for stocks that have high volume. **The annual rate of return** indicates how much the investment grew or shrunk over one year. Overall in 2017, green energy stocks had positive yearly return ratios. In 2018 the situation changed significantly, and the majority performed with negative returns. All except one green-energy stock (_RUN_) indicate risky investments. 
- To refactor VBA code, I added the following steps to the starter code provided in this Challenge (_Steps 1a through 4_). 
![Screenshot_VBA Refactoring](https://user-images.githubusercontent.com/100629325/174422364-82797df0-7f66-4097-9c0f-028974ab45cb.png)
The outputs for the 2017 and 2018 stock analyses in the refactored "All Stock Analysis" script match the output from the original script; however, the code after refactoring runs seven times faster than before. The screenshots below show the execution time of the code before and after refactoring.

![Execution time before refactoring_2017](https://user-images.githubusercontent.com/100629325/174423444-b6e6c40b-248d-4d59-a09a-b65e7addfb90.png)![VBA_Challenge_2017](https://user-images.githubusercontent.com/100629325/174423448-21f982cd-05f7-4a4c-9da5-6fc772bf8861.png)
### Summary
To improve the performance of the original code, we established the output arrays: `tickerVolumes`, `tickerStartingPrice`, and `tickerEndingPrice` in _Step 1b_. These arrays allow us to access the elements using a single variable, `tickerIndex` (_Step 1a_), instead of defining multiple variables. Additionally, it allowed us to avoid creating nested loops and switching between worksheets. The refactored code makes the information faster to retrieve and modify, thus significantly decreasing the macro run time. 
The downside of refactoring code can be spending a lot of time figuring out which specific lines of code need to be addressed. We could end up with new bugs and errors in the code.
