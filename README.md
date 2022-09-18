# Stock Analysis

## Overview of Project

### Purpose
This project aims to provide automated stock data analysis to help the customer decide which tickers are successful and worthy of investing in. The status of the tickers is compared based on two primary measures: total daily volume and the return value. Based on those calculated metrics, investing in tickers with the highest values of total daily volume and return value is recommended. 

## Results
The excution time of the refactored code has been improved by 99.7% (about 0.046 second) over the original code (about 20.42 seconds). The main reason for this improvemnt is becuase we don't nested for loop in the refactored code. In the original code, whenever we start calcuating total daily volume and return metrics, we loop all over the row multiple time for each ticker. While in the refoactored code, we loop through all row only once to calcaute the tickers' perfromnce metrics. In addition, in the refactored code, we save those calcualted matric in an array and then move to the analysis sheet to output those results, however, in the original code, whenevr we calcuate a ticker value we move to the analysis sheet and go back to the data sheet again to calcute the meausres. Which add many extra steps be going back and forth between worksheets for each ticker. The Image below show a comparison between both codes in terms of the main part impacting the excution time. In the refactored code, nested for loop was eliminated, we loop through the stock data worksheet only once to calcaute all tickers information. 

<img width="622" alt="refactored code part1" src="https://user-images.githubusercontent.com/48078471/190880479-296f3d34-94fb-4032-92ef-6b90e6f9cf32.png">


## Summary

1. What are the advantages or disadvantages of refactoring code?

2. How do these pros and cons apply to refactoring the original VBA script?
