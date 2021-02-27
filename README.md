# Stocks-Analysis

## Table of contents
* [Project title](#project-title)
* [Technologies](#technologies)
* [Overview](#overview)
* [Analysis](#analysis)




## Project title
Stocks Analysis - Module 2 Challenge 

## Technologies
Microsoft Excel for Microsoft 365 MSO

## Overview 
Steve’s parents have decided to invest all their money into DAQO New Energy Corporation ($DQ). Steve’s concerned about diversifying their funds and wants to analyze a handful of green energy stocks in addition to DAQO stock. He has created an Excel file and we were asked to analyze this dataset by using an extension to Excel built to automate tasks: Visual Basic for Applications. 
By creating different VBA macro, we were able to find the total daily volume (to measure how actively a stock is traded) and yearly return (to calculate the percentage difference in price from the beginning of the year to the end of the year) for each stock ($DQ, $AY, $CSIQ, $ENPH, $FSLR, $HASI, $JKS, $RUN, $SEDG, $SPWR, $TERP, $VSLR) 

## Analysis
### Results ###
To compare the stock performance between 2017 and 2018, we were instructed to create this sheet (labeled, “All Stocks Analysis”):

![](img/All_Stocks(2017).png)
 

From these charts, one can conclude that the “green energy” sector performed better in year 2017 than in year 2018. All stocks except one ($TERP) had positive yearly return in year 2017. In year 2018, $ENPH and $RUN were the only stocks that performed well, and their total daily volume had increased massively. One can also conclude that $AY is a low volatile stock in this dataset since the stock only moved less than +/-10% in both years. 

When we refactored the original VBA script, the execution time for year 2017 was decreased by about .27 and by .06 for year 2018.

Refactoring code could make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code is not always be the best way to accomplish a task. However, sometimes refactoring code process could change the meaning of the original code without one expecting it or make the code less efficient. 
In this assignment, refactoring the original VBA made the VBA script run faster, which shows that the refactored code was more efficient than the original. The process was relatively easy and beneficial since the coder for the original VBA script and the refactored one are the same, which means the coder understands the purpose and the meaning of the original code. 



