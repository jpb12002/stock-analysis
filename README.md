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
Examining the stock outcomes from 2017 and 2018, we can see that 2017 was a more successful year for the stock market overall (at least for the stocks we are analyzing). Almost every stock in 2017 had a positive return while only two stocks had positive returns in 2018. Because the client Steve is searching for a stock return that is consistently trending upward, we also looked for stocks that had positive returns in both years. There were only two stocks that had positive returns in 2017 and 2018: ENPH and RUN. We can also see that the total daily volume for these stocks dramatically increased from 2017 to 2018. These trends show that investors are confident in the growth of these stocks moving forward. Therefore, our client Steven can confidently recommend to his parents that they should invest in either ENPH or RUN. 

#### 2017 Stock Outcomes
![Image of 2017 stock outcomes](https://github.com/jpb12002/stock-analysis/blob/main/2017_Stocks_Analysis.png)

#### 2018 Stock Outcomes
![Image of 2018 stock outcomes](https://github.com/jpb12002/stock-analysis/blob/main/2018_Stocks_Analysis.png)

## Summary 
- There are several advandatages to refactoring code, the most obvious one from this VBA project is making the execution time of your code faster. By refactoring our original "AllStockAnalysis" code, we saved approximately 0.4 seconds when analyzing the 2017 and 2018 stock data. While the speed increases we saw in this project were not massive, this will become much more important once we are working with larger data sets. Small time savings could translate into minutes or hours of time saved for you and the client. Another advantage with refactoring is that it helps to improve your coding skills and find new coding designs that can make future projects easier. The main example in this VBA project was that avoiding nested loops (when possible) in the refactored code made it run more efficiently. By defining more variables and arrays outside of loops, we can avoid excessive nesting. 

- There were some disadvantages that appeared when refactoring our original code throughout this VBA project. 
