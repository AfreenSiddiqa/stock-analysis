# stock-analysis

## Overview of Project: Explain the purpose of this analysis.

The purpose of these analysis was to help Steve, he wants to look at a different set of stocks in the future. With this in mind, we should create a flexible macro for running multiple stocks. By carefully reusing the code we've already written for DQ, we can write a macro with this flexibility.

Since this code might gets a little complicated, we should start by writing a basic outline of the program flow. Then we'll use comments to organize it all before we write the actual code.

Our new macro should do the following:

Format the output sheet on the "All Stocks Analysis" worksheet.
Initialize an array of all tickers.
Prepare for the analysis of tickers.
Initialize variables for the starting price and ending price.
Activate the data worksheet.
Find the number of rows to loop over.
Loop through the tickers.
Loop through rows in the data.
Find the total volume for the current ticker.
Find the starting price for the current ticker.
Find the ending price for the current ticker.
Output the data for the current ticker.


Steve's parents want to know how actively DQ was traded in 2018. They believe that if a stock is traded often, then the price will accurately reflect the value of the stock. If we sum up all of the daily volume for DQ, we'll have the yearly volume and a rough idea of how often it gets traded.

Steve will probably want to run this analysis for each year, so let's update our code to run for any year, not just 2018.

In this challenge, we will edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that we did in this module. Then, we will determine whether refactoring the  code successfully made the VBA script run faster. 



Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

[VBA_Challenge_2017](https://user-images.githubusercontent.com/93686963/142783542-45ca8e09-6fd1-4e7f-b787-f81fc987b8a2.PNG)

[VBA_Challenge_2018](https://user-images.githubusercontent.com/93686963/142783555-c44dc999-062e-4b55-9af8-f3d01b1b3bef.png)


Summary: In a summary statement, address the following questions.

What are the advantages or disadvantages of refactoring code?
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

How do these pros and cons apply to refactoring the original VBA script?
