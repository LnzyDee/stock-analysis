# VBA Challenge

## Overview of Project

Creating code to run through a large of amount of stock data to compare how all the stocks fared through the years 2017 and 2018.

After coding in the workbook to sort through the dataset, we want to refactor the initial code to see if the new code runs faster.

## Results

### Stocks Results
![2017 Stocks Results](https://github.com/LnzyDee/stock-analysis/blob/main/VBA_Challenge_2017-stocks.png)  ![2018 Stocks Results](https://github.com/LnzyDee/stock-analysis/blob/main/VBA_Challenge_2018-stocks.png)

The stock performances in 2017 were much better than 2018; a significant drop in performance is shown in 2018, with some stocks well over 100% in returns from 2017 to negative percentage returns in 2018.

### Run Time Results Comparison
![2017 Original Code Time Results](https://github.com/LnzyDee/stock-analysis/blob/main/VBA_Challenge_2017-oldcodetime.png)  ![2017 Refactored Time Results](https://github.com/LnzyDee/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)  

![2018 Original Code Time Results](https://github.com/LnzyDee/stock-analysis/blob/main/VBA_Challenge_2018-oldcodetime.png)  ![2018 Refactored Time Results](https://github.com/LnzyDee/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

Above shows the original code run time results on the left, with the refactored results to the right. The refactored script shows over a second increase in speed compared to the original script. The refactored script for both years analysis clocked in under a second to run. The original script for both years analysis took almost a second and a half longer to run than the refactored script.

## Summary

Overall, it seems refactoring the code helped to run the analysis faster by about a second.
- An advantage of refactoring the code is that it did help to run the analysis faster as stated above. A disadvantage of refactoring code means having to do basically the same thing you previously did again but making changes to have more efficient code; sometimes those changes could end up in errors in the code if things aren't entered correctly. It may take more time, but in the end, it can improve overall performance. 
- I definitely faced these pros and cons when refactoring the VBA script. I had a few issues with sorting the arrays and which to assign where; often times I had compile errors come up because of this, as below:
 ![vbascrnshot2](https://user-images.githubusercontent.com/96149719/159102657-996074f9-470b-497b-8d9b-570bb2f60480.png)
 Despite this, it did end up making the code more efficient.
