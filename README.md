# stock-analysis
Performing analysis on green energy stocks to determine total daily volume and yearly return.

## Overview
Steve just finished school with a finance degree. His parents have elected to be his first clients and are specifically interested in investing in DAQO New Energy Corp (DQ). Steve was a little worried and wanted to diversify his parents portfolio while also doing research to determine if the DQ stock was a worthwhile investment. Steve has decided to analyze a handful of green energy stocks in addition to DQ. To make this analysis more efficient, VBA will be used to automate the procedure.

## Results
The results will explain the performance of the stocks in 2017 and 2018 as well as the steps that were taken to make the code used to automate gathering the information for the analysis more computationally efficient.

### Stock Performance
![image_name](https://github.com/Mugunthan24/stock-analysis/blob/main/resources/Stock%20Performance_2017.PNG)
![image_name](https://github.com/Mugunthan24/stock-analysis/blob/main/resources/Stock%20Performance_2018.PNG)

In 2017, all of the stocks with the exception of TerraForm Power (TERP) provided a positive return on investment when looking at the yearly returns. TERP peroformed the worse of all the stocks when looking at the yearly return of -7.2%. Of the 12 green energy stocks, DQ had the greatest return of 199.4%, nearly tripling in value in 2017.

In 2018, all the stocks with the exception of Enphase Energy (ENPH) and Sunrun (RUN) did not provide a positive return on investment when looking at the yearly returns. ENPH had a return of 81.9% and RUN had a return of 84%. Of the 12 green energy stocks, DQ had the lowest a return of -62.6% in 2018.

### Computational Efficiency
![image_name](https://github.com/Mugunthan24/stock-analysis/blob/main/resources/VBA_Challenge_2017%20(Original).PNG)
![image_name](https://github.com/Mugunthan24/stock-analysis/blob/main/resources/VBA_Challenge_2018%20(Original).PNG)

In the original subroutine that created the analysis table for the 12 stocks, the For Loop would iterate through all 3013 rows for each stock (12x), even if the information (volume, starting price, and ending price) was already captured for that stock. The execution time for when the code was run for 2017 and 2018 was approximately 0.95 seconds.

![image_name](https://github.com/Mugunthan24/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)
![image_name](https://github.com/Mugunthan24/stock-analysis/blob/main/resources/VBA_Challenge_2018.PNG)

The refractured code took advantage of the fact that the stock's ticker and dates were sorted in ascending order (A-Z). This makes it possible to iterate through all 3013 rows just once. Instead of iterating through all 3013 rows for each stock, in the ticker column of the data set, once the current ticker value (Cells(i, 1)) we have been looking for does not equal the next ticker value (Cells(i + 1, 1)) then we can move on and start populating information for the next ticker (tickerIndex = tickerIndex + 1). After refractoring the code and implementing this change, the execution time when the code was run in 2017 and 2018 was approximately 0.09 seconds.

Additionally, when the code is run the spreadsheet updates to show each step, although it happens quickly and is hard to notice. For the Stock Analysis subroutine, this increases the execution time slightly. At the beginning of the subroutine, if we add the line "Application.ScreenuUpdating = False" the screen will no longer update while the code is being run and this further reduces the execution time.

## Summary
The summary will provide the advantages and disadvantage of refractoring code in general and in regards to the Stock Analysis subroutines (Original vs. Refractored).

### Advantages and Disadvantages to Refractoring Code in General
The potential advantages to refractoring code in general are:
- reducing execution time so that the code runs faster and is computationally efficient
- allows coder to gain familiarity with current code and discover bugs
- simplifying code makes it easier to update in the future

The potential disadvantages to refractoring code in general are:
- the code can be extremely complicated and take a lot of time to refractor
- the task of refractoring code can be costly
- the person reviewing the code can miss/remove something vital causing bugs that may or may not be caught for a while

### Advantages and Disadvantages to Refractoring Stock Analysis Subroutine
Refractoring the Stock Analysis subroutine has many advantages. These advantages include making the code more computentationally efficient and reducing the run time allowing Steve to get the desired information for his analysis quicker. Furthermore, it allows us to gain familiarity with the code and understand exactly what it is trying to do step by step. In doing this, bugs can be discovered. Lastly, if updates ever need to be made in the future then this can be done more easily since the code has been "cleaned, making it more easily understandable if a new developer comes along and wants to make modifications (e.g. order of data table columns change, so code needs to be updated to account for update).

The disadvantages of refractoring code associated with time consumption and high cost are not applicable to the Stock Analysis subroutine because the code is not very complicated therefore making refractoring likely to be quick and inexpensive. However, the developer refractoring the code whether it is the person who created it or an entirely new person can miss/remove details in the code that can potentially cause bugs. For example, the developer refractoring the code may accidentially change the if condition required to determine starting price.

