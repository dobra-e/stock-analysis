# Green Stocks Analysis
## Project Overview



### Purpose
The purpose of this analysis is to understand the performance of green-energy stocks and determine which are worth investing in. To make this determination, both total daily volume and percent yearly return were analyzed.

## Results
### The Code
The file contains two VBA scripts - AllStocksAnalysis and AllStocksAnalysisRefractored - that have different code, but produce the same output. The differences in the code means one runs more efficiently, and therefore, faster than the other. To run the script, users select one of the "Run Analysis..." buttons and enter either 2017 or 2018 in the pop-up box to produce the output. To start over, users select "Clear Worksheet" and can run a new analysis.

Both the original and refractored code start in the same way. A table for the output was formatted and an array of all tickers was initialized. The refractored code then uses three additional arrays `tickerVolumes` `tickerStartingPrice` and `tickerEndingPrice`, whereas the original code used variables to hold the data. The use of variables required nested `for` loops and switching between worksheets to interate through the stocks. Arrays, on the otherhand, allowed the use of separate `for` loops that didn't require switching between worksheets.

| ![Original Script](/Graphics/OriginalScript.png)|![Refractored Script](/Graphics/RefractoredScript.png)
|:--:|:--:|
|*Original script illustrated the nested `for` loop (purple line) and additional worksheet activation (blue arrow)*|*Refractored Script showing the three separate `for` loops (pink lines)*|






#### Run Times
| ![VBA Script Run Times](/Graphics/2017_original.png)|![VBA Script Run Times](/Graphics/2017_refractored.png)
|:--:|:--:|
|*VBA Script Run Time for 2017 Stocks (Original)*|*VBA Script Run Time for 2017 Stocks (Refractored)*|


| ![VBA Script Run Times](/Graphics/2018_original.png)|![VBA Script Run Times](/Graphics/2018_refractored.png)
|:--:|:--:|
|*VBA Script Run Time for 2018 Stocks (Original)*|*VBA Script Run Time for 2018 Stocks (Refractored)*|


### A Comparison of 2017 & 2018 Stocks
| ![Green-Energy Stock Comparison (2017 & 2018)](/Graphics/StockComparison.png) | 
|:--:| 
| *Green-Energy Stock Comparison (2017 & 2018)* |


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
* Refactoring decreases the time it takes to run a program especially if thousands of stocks are being analyzed.
* Cleaner code
* Refractor again to improve the tickers code

#### Cons
* Time consuming
* If unfamiliar with the code or written by someone else, it may be easier to rewrite the script entirely

