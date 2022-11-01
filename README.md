# Stock Analysis with VBA

## Overview of Project

This project aims to refactor code that analyzes a dataset of stock volume and returns for whatever year is inputted. The goal of refactoring the code is to make it more efficient so that if there were to be more information added to the dataset, the code will still run quickly. 

## Results

### Stock Performance Between 2017 and 2018
In 2017, almost all stocks saw a positive return in a range of 5.5% up to 199.4%. The only stock that had negative returns was TERP, of -7.2%. A year later, ten of these stocks saw a negative return, while only two saw a positive return. Almost all stocks saw a marked decline in their return percentage, with two exceptions. TERP increased its return slightly from -7.2% to -5.0%. RUN increased its return from +5.5% in 2017 to +84.0% in 2018. Along with RUN, ENPH was the only other stock to see a positive return during 2018 (+84.0%), although it was a smaller return than in 2017 (+129.5%.)
    
   ![2017_All_Stocks_Analysis](https://user-images.githubusercontent.com/114192448/199311078-b461fccf-076c-4c18-b4a2-17eb2fab8ad8.png)
   ![2018_All_Stocks_Analysis](https://user-images.githubusercontent.com/114192448/199311121-5379d963-873c-4aff-9192-550b83eadd08.png)

### Refactored Code and Run Time
In order to ensure that the code would remain efficient even if large amounts of information was added to the dataset, it needed to be refactored. in order to refactor the code, the variable tickerIndex was added to the code. This tickerIndex variable assisted with looping through the array of tickers more efficiently, and therefore quickly, as shown in the screenshots below. 

#### 2017 Analysis Before Refactoring
![2017_Unfactored_Time](https://user-images.githubusercontent.com/114192448/199312243-99737be9-7b2b-4d63-8b7e-016077ecfd24.png)

#### 2017 Analysis After Refactoring
![VBA_Challenge_2017](https://user-images.githubusercontent.com/114192448/199312262-a8445345-5123-4650-b6b3-e3e6de1498d2.png)


#### 2018 Analysis Before Refactoring
![2018_Unfactored_Time](https://user-images.githubusercontent.com/114192448/199312274-10f1500e-538b-4f14-8de6-338b69234323.png)


#### 2018 Analysis After Refactoring    
![VBA_Challenge_2018](https://user-images.githubusercontent.com/114192448/199312291-cdf3dc27-6e77-497e-af21-e3887b817ccf.png)


    
## Summary
### 1. What are the advantages or disadvantages of refactoring code?
One of the advantages of refactoring code is to improve the efficiency of the code running the analysis. As noted in the online modules, refactoring code could lead to using less memory, using fewer steps, and making the code easier to read by improving its logic.
A disadvantage of refactoring code is the possibility of introducing bugs and errors into your code, resulting in additional time spent debugging [^1]. Another disadvantage is that there is no standard definition of what a "neat code" is, therefore different programmers could have different ideas of what will be most easily read [^2].

[^1]: https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/ 
[^2]: https://www.ionos.com/digitalguide/websites/web-development/what-is-refactoring/     
    
### 2. How do these pros and cons apply to refactoring the original VBA script?

As stated previously, refactoring this code had a significant impact on the efficiency of the analysis, shown by the decreased time spent running the analysis. If more data were to be introduced to the dataset, this code should continue to run quickly due to the refactoring.
In this project, I found that refactoring this code resulted in a lot of time working on bugs and errors. Refactoring also resulted in more lines of code, which could be argued that would be more confusing and less easily read. 
