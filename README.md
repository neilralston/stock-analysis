# VBA of Wall Street (Module 2)

[Link to VBA_Challenge Excel File](VBA_Challenge.xlsm)

## Overview of Project

### Purpose
Assessed how a set of twelve "green stocks" performed over a given year in terms of total volume traded (Total Daily Volume) and annual return (Return).

Used VBA to develop a macro that calculates Total Daily Volume and Return.  

* To accomplish this task, the macro uses a combination of conditionals, arrays, loops, nested loops, and indexes.

* The macro uses conditional formatting for visual formatting, along with custom buttons for an easy user interface.

* Finally, the macro was refactored so that it could be reused to assess a much larger dataset and still run efficiently.

[Link to Refactored Macro Script](VBA_Challenge_macro_text.txt)

## Results

### Stock Performance Between 2017 and 2018

In 2017, all but one of the twelve stocks had a positive return.  The best performer was "DQ", followed by "SEDG".

![T2017 Stock Performance Chart](2017_results.PNG)

In 2018, only two of the twelve stocks had a positive return -- "ENPH" and "RUN".  

![T2018 Stock Performance Chart](2018_results.PNG)

The top performers from 2017 both posted negative returns in 2018.  This suggests a high level of market volatility in this sector that should be considered before investing.

### Execution Times Between Original and Refactored Scripts

The original macro scripts run in approximately one second.

![Original Script Execution Time 2017](2017_time.PNG)

![Original Script Execution Time 2018](2018_time.PNG)

The refactored macro scripts run in approximately 0.20 seconds (about one-fifth of a second).

![Refactored Script Execution Time 2017](VBA_Challenge_2017.PNG)

![Refactored Script Execution Time 2018](VBA_Challenge_2018.png)

## Summary

What are the advantages or disadvantages of refactoring code?
* A primary advantage of refactoring is to produce higher quality code that runs more efficiently.

* A disadvantage is that the refactoring process makes code more complex, requiring more functionality testing and creating greater opportunities for "breaks".

How do these pros and cons apply to refactoring the original VBA script?
* The code runs faster after refactoring, which will make it more efficient if assessing a larger dataset.

* I ran into trouble writing the correct syntax for the additional indexing arrays required to refactor the code. Fortunately, I was able to find help from an online source that displayed code written to solve a similar problem.  This helped me to better understand the correct syntax for the assignment. https://github.com/caseychen3605/stock-analysis 



