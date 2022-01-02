# stock-analysis
Performing analysis on green energy stocks to determine total daily volume and yearly return.

## Overview
Steve just finished school with a finance degree. His parents have elected to be his first clients and are specifically interested in investing in DAQO New Energy Corp (DQ). Steve was a little worried and wanted to diversify his parents portfolio and while also doing research to determine if the DQ stock was a worthwhile investment. Steve has decided to analyze a handful of green energy stocks in addition to DQ. To make this analysis more efficient VBA will be used to automate the procedure.

## Results
The results will explain the performance of the stocks in 2017 and 2018 as well as the steps that were taken to make the code used to automate gathering the information for the analysis more computationally efficient.

### Stock Performance
In 2017, all of the stocks with the exception of TerraForm Power (TERP) provided a return on investment when looking at the yearly returns. TERP peroformed the worse of all the stocks when looking at the yearly return of -7.2%. Of the 12 green energy stocks, DQ had the greatest return of 199.4%, nearly tripling in value in 2017.

In 2018, the performance of all the stocks with the exception of Enphase Energy (ENPH) and Sunrun (RUN) did not provide a return on investment when looking at the yearly returns. ENPH has a return of 81.9% and RUN had a return of 84%. Of the 12 green energy stocks, DQ had the lowest a return of -62.6% in 2018.

### Computational Efficiency
In the original subroutine that created the analysis table for the 12 stocks, the For Loop would iterate through all 3013 rows for each stock (12x), even if the information (volume, starting price, and ending price) were already captured for that stock. The execution time for when the code was run for the 2017 data was 0.94 seconds. 

The refractured code took advantage of the fact that the stock's ticker and dates were sorted in ascending order (A-Z). This makes it possible to iterate through all 3013 just once. Instead of iterating through all 3013 rows for each stock, in the ticker column of the data set, once the current ticker value (Cells(i, 1)) we have been looking for does not equal the next ticker value (Cells(i + 1, 1)) then we can move on and start populating information for the next ticker.
