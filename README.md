# Stock Analysis

## Overview of Project

### Purpose
This project aims to provide automated stock data analysis to help the customer decide which tickers are successful and worthy of investing in. The status of the tickers is compared based on two primary measures: total daily volume and the return value. Based on those calculated metrics, investing in tickers with the highest values of total daily volume and return value is recommended. 

## Results
The execution time of the refactored code has been improved by 99.7% (about 0.046 seconds) over the original code (about 20.42 seconds). The main reason for this improvement is that we do not nest for loop in the refactored code. In the original code, whenever we start calculating total daily volume and return metrics, we loop all over the row multiple times for each ticker. While in the refactored code, we loop through all rows only once to calculate the tickers' performance metrics. In addition, in the refactored code, we save those calculated metrics in an array and then move back to the analysis sheet to output those results. However, in the original code, when we calculate a ticker value, we move to the analysis sheet and back to the datasheet again to calculate the measures, which adds many extra steps by going back and forth between worksheets for each ticker. The image below compares both codes regarding the main part impacting the execution time. 

In the refactored code, the nested for loop was eliminated; we loop through the stock data worksheet only once to calculate all ticker information. 
<img width="622" alt="refactored code part1" src="https://user-images.githubusercontent.com/48078471/190880479-296f3d34-94fb-4032-92ef-6b90e6f9cf32.png">

In addition, the tickers' perfromnce measures are stored in arrays and then we output them into the analysis worksheet in one step using a for loop as in the image below. 
<img width="620" alt="refcatored 2 part2" src="https://user-images.githubusercontent.com/48078471/190880573-a36bd382-a64d-4b8d-9c50-8bbfb64bb896.png">
The execution time of the refactored scrpit is 0.046 as shown in the following image. 
<img width="294" alt="execution time of refactored code" src="https://user-images.githubusercontent.com/48078471/190880583-2ec75324-6906-4119-882a-d912098f9ad7.png">

The original code took 20.42 seconds to run because of the nested for loop, where to calculate each ticker's values, the code loops through all the rows. Besides, the code keeps moving between data & analysis sheets whenever a ticker's values are calculated. This step is redundant, and emanating it in the refactored code helps to increase efficiency. Images below show the execution time, and the main part causes the long run time. 
<img width="556" alt="original code" src="https://user-images.githubusercontent.com/48078471/190880644-85ad88cf-9f53-436c-8a2e-193e8b0db74c.png">
<img width="281" alt="execution time of original code" src="https://user-images.githubusercontent.com/48078471/190880645-d91e5edd-e202-44d0-964b-16f543edff56.png">


## Summary

1. What are the advantages or disadvantages of refactoring code?\
The main advantage of the refactored code is the efficiency, and it takes considerably less time to run the same functionality of the original script. The disadvantage is that I found the refactored code's logic harder to follow than the original code. 

2. How do these pros and cons apply to refactoring the original VBA script?\
The original code logic is easy to read, but it took a long time to run since we kept moving between worksheets to calculate values and then output them in another sheet. It also keeps looping through the rows even though we found the information of interest for a particular ticker, which is not adding any value. While the refactored code removes those redundant steps, for example, we loop through the data only once to calculate measures of all tickers, and we output the data in a single step to the analysis sheet to save time moving back and forth between sheets. 
