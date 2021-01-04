# VBA Challenge

## Overview of Project

Use VBA in Excel to analyze Wall Street Stock information from the excel file “Green Stocks" dataset. 

### Purpose

The purpose of this project was to create an easy to use macro application that allowed users with limited familiarity with excel to be able to perform analysis on the given 
dataset with regards to a ticker's total daily volume and return on investment on yearly increments. 

The macros which were also made as buttons to allow for ease of use for end user, would run through a tickers daily opening and closing price and provide a summary  for the 
total daily volume, and return on investment. This summary would be created as a new table on a new sheet to filter the necessary and pertinent information the end user may 
want. The macro would also highlight positive returns in green and negative returns in red to give insight for better decision making for the end user to make a continued 
investment decision on a particular stock.

## Analysis and Challenges

Analysis:
The analysis for this project depends on the window for investment, on a yearly basis one can determine if a stock would be a good short term buy or sell. Since the 
daily volume is being tracked an investor can determine based on volume if a stock was gaining momentum and would be a good short term buy. Volume analysis would also allow 
such an investor to determine if a stocks trend is viable or weak and may trend in the wrong direction. 

A 2017 analysis for instance indicates 11 of the 12 stocks yielded excellent returns.
 

Any reasonable investor would be ecstatic on this return and more likely than not stay invested in such stocks. However doing so would be detrimental based on 2018’s return. 

The investor that kept their investment the same in 2017 or increased their positions would have lost a fair sum of money. 
As the application is currently made as long the stock data were dynamic and continuously update further analysis could be made and additional macro could be made to automate a power query to update this data regularly.  The possibility of expanding this application also comes with the possibility the analysis can slow down so stream lining the macro would be need to speed up execution time. As the macro is constantly run the execution time does speed up from the initial run. 

2017 Run Time Before:
 
.328125 seconds

2017 Run Time After:
 
.046875 seconds

2018 Run Time Before:
 
.328125 seconds

2018 Run Time After:

.046875 seconds

The potential of this application can track any stock provided the data for that stock is loaded allowing this application to versatile in providing a great deal of analysis without much need to altering the macros that run the analysis.

Challenges: 
VBA is extremely particular on syntax and if there is an error in misspelling or a variable is just written slightly wrong the code will not run or worse the code does not run and the debug does not catch the slight error. VBA has been in use quite a while so there is a plethora of documentation available the robustness of VBA is simple not comparable to other languages. The end user would be greatly limited in terms of what VBA allows to be done within the language framework. 

## Summary

1.	What are the advantages or disadvantages of refactoring code?

The main advantage is the code is simplified and much easier to maintain. As code is streamlined one get or create cleaner code. The greatest disadvantage for refactoring would be the higher risk of bugs in the code and even more time trying to debug the code to work for the intended needs. This also leads to the chance of having to not have any documentation available to help one use a cleaner line of code. 

2.	How do these pros and cons apply to refactoring the original VBA script?

Refactoring allowed for the removal of the nested loops which made the code much easier to read and the number of lines needed to run the code reduced significantly. This also allowed for more robustness to check other stocks in the future without having to do much more coding. The downside was knowing how the indexing would behave and impact the rest of the code. This lead to quite of bit of debugging for the loops that were left or needed to be added back if  the line as commented out. 
