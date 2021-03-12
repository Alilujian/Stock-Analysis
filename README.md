# Stock-Analysis
Stock analysis using VBA

## Overview of Project
Steve and his parents are about to invest in green energy stocks. In this project, we will using VBA tool to analysis the stocks data and determine their daily volume and yearly returns in 2017 and 2018. Throughout this project, we will compare the refactored code and original code to see if the refactored code improve the code performance. By completing this project, I would learn the use of arrays, nested loop, and conditional formarting in VBA script.

## Results
### Stock Analysis Results 
In 2017, except for TERP, all other stocks had a positive returns. Among the 12 stocks, DQ achieved the highest yearly returns with 199.4%. Where RUN gained the lowest positive return with 5.5%.\
In 2018, except for ENPH and RUN, all toher stocks had a negative yearly returns. ENPH and RUN achieved similar yearly returns around 80%. And both of their daily volume increased by more than $200,000,000 from 2017. JKS reached a highest negative returns of -60.5% among the 12 stocks, and its daily volume also drops.\

![VBA_Challenge_2017](VBA_Challenge_2017.png)\
Figure 1: the daily volume and yearly returns of all stocks in 2017.

![VBA_Challenge_2018](VBA_Challenge_2018.png)\
Figure 2: the daily volume and yearly returns of all stocks in 2018.

### Original Code Vs Refactored Code


![original_code](original_code.png)
Figure 3: the original code.

![Refactor_code](Refactor_code.png)
Figure 4: the Refactored code.

#### Advantages
1. Reducing code run time

#### Disadvantages
1. Time consuming



## Summary
In conclusion, only the ticker ENPH and RUN will be considered as a good investment because of their positive returns in 2018. Even though the green enrgy stock DQ increased its daily volume, but it was not enough to cover the lost. So that the 2018 returns for DQ was negative.\
As for code refactoring, ther are both advantages and disadvantages. In my opinon, code refactoring is necessary because it can highly improve the readability and performance of our code.
