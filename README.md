# VBA of Wall Street

## Overview of Project

### Purpose
The purpose of this project was to increase the speed at which our client Steve can analyze the entire stock market. Our previously submitted workbook enabled Steve to analyze a small subset of the stock market with relative efficiency. However, Steve may not have time to wait for our current code block to run through every possible stock option. Therefore, our goal for this project is to try and increase the speed of our macro by refactoring our code and comparing the time it takes to execute the new macro vs. the old macro. 

## Results

### Refactored Code Comparison
Included here are examples of both the original macro code, "AllStocksAnalysis", and the refactored macro code, "AllStocksAnalysisRefactored", that can be found in the "VBA_Challenge" workbook. This first code block image is taken from the original "AllStocksAnalysis" macro and shows the nested For loop structure that was created. This nested loop would search through every row in either the "2017" or "2018" data sheet that is chosen by the user to find the rows that included the ticker name at index 0 of the "tickers" array. If the correct ticker name was included in that row, the corresponding data would then be added to the "All Stocks Analysis" sheet to provide values for the total daily volume and net stock return. This process was then repeated eleven more times for every stock ticker in the "tickers" array. 

#### Original Code Block
![Image of original code block](https://github.com/jpb12002/stock-analysis/blob/main/Original_Code.png)

This second code block image is taken from the refactored "AllStocksAnalysisRefactored" macro and shows how we used a counter variable "tickerIndex" to remove the need for a nested For loop structure. By defining a counter variable as well as the three output arrays outside of the For loop, the macro will only have to search through every row in the "2017" or "2018" data sheet one time. The macro searches through every row that includes the ticker name at index 0 of the "tickers" array and once the final row for that ticker is found, the "tickerIndex" counter variable value is increased by 1. The For loop can then continue to search through the remaining rows for the ticker name at index 1 of the "tickers" array. Since the macro only has to search through every row in the data sheet one time, this should significantly reduce the execution time of the macro. 

#### Refactored Code Block
![Image of refactored code block](https://github.com/jpb12002/stock-analysis/blob/main/Refactored_Code.png)

### Original Macro vs. Refactored Macro Execution Time Comparison
The next series of images provides evidence that the refactored macro is more efficient than the original macro at analyzing a large set of stock data. Comparing the times for the "2017" data sheet, we can see that the refactored macro executed approximately 0.4 seconds faster than the original macro. 

#### Original Macro Execution Time ("2017" Data)
![Image of execution time original script 2017 stocks](https://github.com/jpb12002/stock-analysis/blob/main/Original_Analysis_2017.png)

#### Refactored Macro Execution Time ("2017" Data)
![Image of execution time refactored script 2017 stocks](https://github.com/jpb12002/stock-analysis/blob/main/VBA_Challenge_2017.png)

Comparing the times for the "2018" data sheet, we can again see that the refactored macro executed faster than the original macro by approximately 0.45 seconds. We can therefore say with confidence that the refactored macro is more efficient. 

#### Original Macro Execution Time ("2018" Data)
![Image of execution time original script 2018 stocks](https://github.com/jpb12002/stock-analysis/blob/main/Original_Analysis_2018.png)

#### Refactored Macro Execution Time ("2018" Data)
![Image of execution time refactored script 2018 stocks](https://github.com/jpb12002/stock-analysis/blob/main/VBA_Challenge_2018.png)

### Analysis of Stock Outcomes

#### 2017 Stock Outcomes
![Image of 2017 stock outcomes](https://github.com/jpb12002/stock-analysis/blob/main/2017_Stocks_Analysis.png)

#### 2018 Stock Outcomes
![Image of 2018 stock outcomes](https://github.com/jpb12002/stock-analysis/blob/main/2018_Stocks_Analysis.png)

## Summary 
