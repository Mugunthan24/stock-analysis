# stock-analysis
Performing analysis on green energy stocks to determine total daily volume and yearly return.

## Overview
Steve just finished school with a finance degree. His parents have elected to be his first clients and are specifically interested in investing in DAQO New Energy Corp (DQ). Steve was a little worried and wanted to diversify his parents portfolio while also doing research to determine if the DQ stock was a worthwhile investment. Steve has decided to analyze a handful of green energy stocks in addition to DQ. To make this analysis more efficient, VBA will be used to automate the procedure.

## Results
The results will explain the performance of the stocks in 2017 and 2018 as well as the steps that were taken to make the code used to automate gathering the information for the analysis more computationally efficient.

### Stock Performance
In 2017, all of the stocks with the exception of TerraForm Power (TERP) provided a positive return on investment when looking at the yearly returns. TERP peroformed the worse of all the stocks when looking at the yearly return of -7.2%. Of the 12 green energy stocks, DQ had the greatest return of 199.4%, nearly tripling in value in 2017.

In 2018, all the stocks with the exception of Enphase Energy (ENPH) and Sunrun (RUN) did not provide a positive return on investment when looking at the yearly returns. ENPH has a return of 81.9% and RUN had a return of 84%. Of the 12 green energy stocks, DQ had the lowest a return of -62.6% in 2018.

### Computational Efficiency
In the original subroutine that created the analysis table for the 12 stocks, the For Loop would iterate through all 3013 rows for each stock (12x), even if the information (volume, starting price, and ending price) were already captured for that stock. The execution time for when the code was run for 2017 was 0.94 seconds. 

The refractured code took advantage of the fact that the stock's ticker and dates were sorted in ascending order (A-Z). This makes it possible to iterate through all 3013 just once. Instead of iterating through all 3013 rows for each stock, in the ticker column of the data set once the current ticker value (Cells(i, 1)) we have been looking for does not equal the next ticker value (Cells(i + 1, 1)) then we can move on and start populating information for the next ticker (tickerIndex = tickerIndex + 1). After refractoring the code and implementing this change, the execution time when the code was run for 2017 was 0.21 seconds.

Additionally, when the code is run the screen updates to show each step, although it happens quickly and is hard to notice. For the Stock Analysis subroutine, this increases the execution time slightly. At the beginning of the subroutine, if we add the line "Application.ScreenuUpdating = False" the screen will no longer update while the code is being run and this further reduces the execution time.

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
- the person reviewing the code can miss something vital causing bugs that may or may not be caught for a while

### Advantages and Disadvantages to Refractoring Stock Analysis Subroutine
