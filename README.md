# Stock-Analysis
An excerise in refactoring a VBA script used for performance analysis of a selection of green stocks. 

## Overview of Project
### Project Context
The very real and not at all made up recent finance graduate, Steve, is eager to help his folks determine how to best invest in green stocks. Their current strategy is to select a stock that is meaningful to them, like DAQO New Energy Corp, a stock that reminds them of how they met at Dairy Queen (ticker: DQ) so many years ago. Steve doesn't recall this particular strategy being covered in any of his classes and would prefer to approach stock selection with an analytical approach, reviewing 2017 and 2018 performance of several green stocks to evaluate which stocks, if any, may produce better return for his parental unit, especially as they will determine his potential future inheritance. 

### Purpose of Analysis
After building the VBA scripts in many steps, initially assessing a single stock, Steve has asked us to refactor our code to allow for expanded datasets that may include the entire stock market and analysis of multiple prior years' performances. Because the original code ran multiple times through the dataset, this approach may be prohibitively time consuming on a larger dataset. This project provides a refactored version of our prior VBA scripts with a goal to reduce processing time while making the code easy to be understood.

## Results
### Stock Assessment
Using our refactored code to analyze a selection of stocks, we see that the DQ stock Steve's parents selected is among the most volatile. In the images of our output, you can see that while the stock was the highest performer in 2017, it experienced the third highest loss in 2018. 

![Image of 2017 Stock Output](https://github.com/ozloty06/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

![Image of 2018 Stock Output](https://github.com/ozloty06/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

Generally, most of the stocks performed well in 2017 and all but two stocks lost value in 2018. Based on a glance of both 2017 and 2018 performance, Steve may consider recommending the stock with the ticker ENPH to his parents as it was a high performer both years. That said, Steve's classes hopefully taught him the value of diversification and he may be wise to steer his parents away from over investing in a single stock.

### Refactored Code Assessment
While the code ran very quickly in our initial version, the dataset was a small selection of 12 stocks. Our refactored code is intended to be used to analyze a broader number of stocks, potentially hundreds if the span of the analysis goes beyond green stocks. 

As such, even marginal gains of 0.01 sec gains seen in the refactored code for both years are beneficial. More importantly, as this script is likely to be handed off to Steve long-term and possibly not modified more than a couple times a year, the refactoring is most helpful in providing comments and making the code easier for Steve to expand as additional tickers are added. 

## Summary
### Advantages and Disadvantages of Refactoring
Refactoring our code will provide Steve with the advantage of easier code maintenance and make it easy for anyone to read and understand the code, especially after returning to the code at a later date. 

The downside is that refactoring the code took almost as long to complete as our initial code writing required. While refactoring, I made several errors and slightly modified variable names that had to be modified in every instance. 

### Application to Refactoring Stock Analysis VBA Script

Benefiting from taking the time to refactor our script, our code is now clean, well-organized, easy to change and easy to understand, enhancing our ability to maintain this script over time as stock analysis are performed a few times a year. As intended, this code did run more efficiently with the enhancements provided by refactoring.

The disadvantages of taking time and debugging errors was evident in this effort. Code was copied from an original script to help create this version and some of the variable names had to change. Overlooking a single character in a single instance for that variable led to several hours spent debugging. This would be avoided had I had the foresight to do a global find/replace. Live and Learn :) 

Original Variables 
```
If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
  startingPrice = Cells(i, 6).Value
End If
```

Refactored Variables 
```
If Cells(i - 1, 1).Value <> tickerIndex And Cells(i, 1).Value = tickerIndex Then
  tickerStartingPrices = Cells(i, 6).Value
End If
```

Note the pural form of "price" in the refactored version of the code. Keeping code refactoring top of mind as you write code can help avoid challenges later as you are closer to your code when you first write it and even a week away from it can make it easy to forget what was intended. Refactoring may be best done early and often.
