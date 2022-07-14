# Greenstock Analysis Using VBA
Performing analysis on stocks data from 2017 & 2018

## Overview of Project
The purpose of this project was to analyze stock data using Visual Basic Application (VBA)to assess stock performance between the years 2017 and 2018 using total daily volume and percent return. 

## Results

### Comparison of Stock Performance Between 2017 & 2018
Analysis was performed by refactoring and completing the code provided for the greenstocks analysis dataset. Loops were generated for the 12 stock types: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR to assess both the total daily volume and return percentage for each stock type. Output data is organized in alphabetical order and formatted to show positive return percentages in green, neutral with no fill color, and negative returns in red. Stock performance for all stock types, except TERP, performed better in 2017 than in 2018 as seen by increased return percentage (Fig. 1-2). While return percentage showed as a strong indicator of stock performance, total daily volume did not have a correlation with stock performance in 2017 (Fig. 3). In 2018, there was a general positive correlation between total daily volume and percentage return, but this is most likely skewed by high outliers (Fig. 4).

#### Fig. 1
![stock_performance_2017](https://user-images.githubusercontent.com/108199140/178884268-e4bbf64d-0436-474e-b28e-a8ec5abb0907.PNG)

#### Fig. 2
![code_performance_2018](https://user-images.githubusercontent.com/108199140/178884281-e350682b-76fc-478c-a4f9-bc9ce3e217b2.PNG)

#### Fig. 3
![2017_stock_graph](https://user-images.githubusercontent.com/108199140/178884334-8a4a7f24-a9e6-4e27-9f79-cce6c021e41e.PNG)

#### Fig. 4
![2018_stock_graphs](https://user-images.githubusercontent.com/108199140/178884364-e50d062a-26ea-44cc-90a3-8567f4880337.PNG)

The analysis code also determined the run time for the code to run for each performed analysis. Using our system the code for both analyses took the same amount of time to run, 0.0859 seconds (Figs. 5-6).

#### Fig. 5
![2017_code_performance](https://user-images.githubusercontent.com/108199140/178884749-eae1e303-9bdd-4d3e-9555-7f0a5f345b72.PNG)

#### Fig. 6
![2018_code_performance](https://user-images.githubusercontent.com/108199140/178884768-914ae71a-529d-4775-bbcc-8b8f9a8b6194.PNG)

## Summary

### Advantages and disadvantages of refactoring code
The advantages of refactoring code are that you can improve code performance, speed, and cleanliness without changing its function. This allows consistent improvements by multiple users to continue to improve code design and functionality by modifying code instead of rewriting overall saving time. The cons to refactoring code could be the accidental addition of bugs as well as adding errors that could render the code non-functional. 

### Advantages and disadvantages of the refactored VBA code
The advantage of refactoring the VBA code for this analysis is the code performance. Running the base code just to do a single analysis on the ticker "DQ" took much longer compared to the refactored code performing the analysis on all tickers (Fig. 7). The disadvantage of refactoring this code is that it was time consuming to scan through the code and look for errors and ways to adjust code that was already there.

#### Fig. 7) Runtime of DQ analysis code
![unfactored_code_performance](https://user-images.githubusercontent.com/108199140/178886269-ba67f331-e46e-42eb-a57c-5709602e7c33.PNG)


