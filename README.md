# Green Stocks Analysis
## Project Overview
Green energy investors put money into companies that use naturally generated energy, particularly renewable energy sources. Data from twelve green energy stocks in 2017 and 2018 are included in the dataset. One of the included stocks is DAQO, which the client would like to invest in. 

### Purpose
The purpose of this analysis is to understand the performance of DAQO and other green-energy stocks to determine which are worth investing in. To make this determination, both total daily volume and percent yearly return were analyzed.

## Results
### The Code
The file contains two VBA scripts - AllStocksAnalysis and AllStocksAnalysisRefractored - that have different code, but produce the same output. The differences in the code means one runs more efficiently, and therefore, faster than the other. To run the script, users select one of the "Run Analysis..." buttons and enter either 2017 or 2018 in the pop-up box to produce the output. To start over, users select "Clear Worksheet" and can run a new analysis.

Both the original and refractored code start in the same way. A table for the output was formatted and an array of all tickers was initialized. The refractored code then uses three additional arrays `tickerVolumes` `tickerStartingPrice` and `tickerEndingPrice`, whereas the original code used variables to hold the data. The use of variables required nested `for` loops and switching between worksheets to interate through the stocks. Arrays, on the otherhand, allowed the use of separate `for` loops that didn't require switching between worksheets.

| ![Original Script](/Graphics/OriginalScript.png)|![Refractored Script](/Graphics/RefractoredScript.png)
|:--:|:--:|
|*Original script illustrating the nested `for` loop (purple line) and additional worksheet activation (blue arrow).*|*Refractored script showing the three separate `for` loops (pink lines).*|

#### Run Times
The refractored script runs almost four times faster. 

| ![VBA Script Run Times](/Graphics/2017_original.png)|![VBA Script Run Times](/Graphics/2017_refractored.png)
|:--:|:--:|
|*VBA Script Run Time for 2017 Stocks (Original)*|*VBA Script Run Time for 2017 Stocks (Refractored)*|


| ![VBA Script Run Times](/Graphics/2018_original.png)|![VBA Script Run Times](/Graphics/2018_refractored.png)
|:--:|:--:|
|*VBA Script Run Time for 2018 Stocks (Original)*|*VBA Script Run Time for 2018 Stocks (Refractored)*|


### A Comparison of 2017 & 2018 Stocks
The output of the analysis is below.

| ![Green-Energy Stock Comparison (2017 & 2018)](/Graphics/StockComparison.png) | 
|:--:| 
| *Green-Energy Stock Comparison (2017 & 2018)* |

#### Yearly Return
From the tables it is clear that green-energy stocks performed better in 2017; all but one stock had a positive yearly return. The average yearly return for 2017 and 2018 was 67.3% and -8.5%, respectively. The stocks that had a positive return both years were ENPH and RUN. These stocks seem to be the best choice for investing. 

DAQO had the highest yearly return (199.4%) in 2017, but a negative return (-62.6%) in 2018. However, across years this would have been a positive investment overall with a 136.8% return.

#### Total Daily Volume
There doesn't appear to be a strong relationship between the total daily volume and return. When sorted by daily volume, the returns don't trend positively or negatively. In 2017, DAQO had the lowest daily volume, but the highest return. In 2018, the stock with the highest daily volume (ENPH) had one of the highest returns. 

## Summary
### Advantages and Disadvantages of Refactoring Code
Refactoring is updating the code to improve the design and/or structure without changing the functionality of the code.

#### Advantages of Refractoring
* Refactoring can make the code more readable and easier to understand. This is especially helpful for future developers who may need to work with the code.
- When the code is cleaner, it is easier and less costly to maintain and to add additional functionality. 
- Clean code is also easier to debug when problems arise.
* The performance of the code may improve, such as decreased run time. With fewer redundancies or complexities within the code, the computer doesn't need to work as hard to execute the code and therefore can provide output faster.

#### Disadvantages of Refactoring
* Refactoring doesn't add any additional functionality. Time is spent updating the code with no immediate product improvement.
* The process of refractoring can take a lot of time, especially if the original script was not written by the individual tasked with refractoring. Without comments, it may be challenging and time-consuming to decifer what the code is doing.
* While refactoring, the code may break leading to the product being unusable for a period of time.

### Pros and Cons of Refractoring Original VBA Script
#### Pros
* The refactored script is more efficient and runs almost four times faster than the original script. This is a huge benefit if the code were to be applied to data with hundreds or thousands of stocks to iterate through. 
* The script is also cleaner and avoids nested `for` loops and switching between sheets. Not only does this speed up the execution of the script, it also makes debugging and adding additional features easier.

#### Cons
* One of the biggest cons in both the original and refractored script is having to manually assign the ticker value in the array. This is not a scalable solution for data with more stocks included. The script could greatly be improved by refactoring to automate the assignment of ticker values.

