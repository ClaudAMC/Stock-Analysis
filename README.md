# Stock-Analysis
---
## Overview of Project
Steve would like to help his parents with their investments. We created a workbook for him that would help with his research, however, he wants to expand the dataset to multiple years. Our original workbook may work for smaller datasets but if the dataset is expanded it may take a long time to execute. 
### Purpose
The purpose of this analysis is to refactor the origional VBA script in order to make it run more efficiently and decrease its execution time. Using this script the stock performance for 2017 and 2018 were evaluated and compared.

## Analysis and Refactoring

The file and scripts; original and refactored can be found in the file below:

[VBA_Challenge.xlsm](ClaudAMC/Stock-Analysis/VBA_Challenge.xlsm)

### Diferences in the code

The  original code we made for Steve looped through an array representing each of the stocks to be analyzed, within that it looped through all of the rows in the worksheet and found the information for the correct variables then displayed it. The variables we were looking for were the total volume of stocks, the starting price for that year, and the ending price for that year.

The refactored code did not have nested loops instead it had 3 seperate loops, additionally, as opposed to one array there were 4: a) stocks, b) volume of stocks, c) starting price, and d) ending price. The first loop initialized the volume of stocks to zero each time. The next one looped through each row and calculated the volume, start and end price. The final loop output the calculated variables for each stock and recorded it in the specified worksheet. Within the second loop an index that was externally set to zero was used to keep reference of which stock the loop iteration was looking for; this index was then increase after each loop iteration.

These changes in the refactored allowed the analysis to run more efficiently and there by decreasing the execution time.

### Analysis of 2017 Stocks: Original vs Refactored

The original code results in a slower execution time as seen in the image below.

![Original 2017 run time](https://user-images.githubusercontent.com/103139895/169166551-b8626045-9956-4e73-b086-13c1c1c6a273.png)

The refactored code results in a faster exection time as seen in the image below.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/103139895/169166375-4c8a2b78-1c33-47ce-9242-e1c7ec1e1f0d.png)

### Analysis of 2018 Stocks: Original vs Refactored

The original code results in a slower execution time as seen in the image below.

![Original 2018 run time](https://user-images.githubusercontent.com/103139895/169167693-ad449794-b92a-4217-93e4-a71067405c42.png)

The refactored code results in a faster execution time as seen in the image below.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/103139895/169167707-73a58fd9-e3a6-4a3a-8b9c-171cb3ea2261.png)

### Comparing 2017 to 2018 Stocks

By looking at the images above we can also compare the return for each stock in each year and how they differed. We can see that in 2017 most of the stocks had a positive return. In fact all stocks except TERP had a positive return with stocks DQ, ENPH, FSLR, and SEDG all having more than a 100% return. This did not stay true for the year 2018. In 2018 only two stocks had a positive return: ENPH, and RUN. All other stocks had a negative return. From this we can draw out the conclusion that the stock ENPH may be worth investing in since it had high positive returns for two years in a row: 129.5% in 2017 and 81.9% in 2018. The RUN stock may also be a good investment however, the magnitude of this positive return may not be as certain as the return was only 5.5% in 2017 but increased to 84% in 2018. This may indicate a positive trend for this stock however, it may also mean the stock is more volatile. As a preliminary conclusion I would  suggest that steve look at investing into ENPH or RUN with his parents however, if possible, I would sugest looking at more years to be able to draw a more accurate trend for both of these stocks.

