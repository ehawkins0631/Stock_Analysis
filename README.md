Overview of Analysis

An Excel workbook containing stock data from years 2017 and 2018 was provided for the analysis.
A client Overview of Analysis

An Excel workbook containing stock data from years 2017 and 2018 was provided for the analysis. 
A client wanted to help his parents diversify their portfolio, which was only invested in DAQO green energy stock.

The goal of the project was to analyze stock performance data using VBA. An existing code was refactored in order to evaluate
if it ran better or faster.

Results

By the end of the analysis, important information was extracted, such as total volume and yearly return values. 
The worksheets were formatted for easier visualization and tools for easy access within the sheet were created.

The analysis showed no significant difference in the outcome as a result of refactoring the code.

Timer 2017 before refactor 0.2973633
Timer 2017 after refactor 0.8081055
Timer 2018 before refactor 3.271484
Timer 2018 after refactor 0.8330078

As shown above, the 2017 timer ran faster before the changes, and 2018 timer ran faster after the changes.

Figure 1:

![image](https://user-images.githubusercontent.com/101227930/188049163-a15dd642-a501-40c9-8445-3cdd02e3349a.png)


Figure 2:

![image](https://user-images.githubusercontent.com/101227930/188049047-d5cfbd99-4d74-4fda-a2db-64321e153eda.png)

Figure 3:
![image](https://user-images.githubusercontent.com/101227930/188049244-fa434f16-f748-4166-976d-a2d7c73bdf7c.png)




In conclusion, ENPH and RUN were the best performing stocks in 2018 (shown in green). DQ was down 63%.

In 2017, however, all of the stocks did extremely well, except TERP (shown in red). DQ had a massive return and SEDG as well; close to 200%.
FSLR and ENPH more than 100%

Summary

Refactoring code is essentially editing the existing code to determine if it will run faster. In order for it to properly
work, it needs to be done in small steps. The steps are completed without major changes to existing code (e.g., external 
behavior, functionality).

Refactoring codes can have its advantages as well as disadvantages. The following are the advantages and disadvantages that
caught my attention:

Disadvantages

Some parts of the code may be repeated in several parts of the code. It may inadvertently introduced unwanted bugs
The process may affect the testing outcomes

Advantages

It is mostly safe since one may not refactor without testing
Since run times may be shorter, it can boost the systems performance
It is possible to make it easier to read than original version




email:  ehawkins0631@gmail.com

twitter: @evahawkins0630

https://www.linkedin.com/in/eva-hawkins-a9b333147/

