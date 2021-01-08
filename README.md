# OVERVIEW: STOCK ANALYSIS USING VBA

## Purpose
#### The purpose of this project to is edit or refactor the Stock Market data set using the VBA solution code to loop through all the data one time in order to collect an entire dataset. We'll determine whether refactoring the code successfully made the VBA script run faster. Finally, make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Background of the project
#### Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

#### In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

#### Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

## Requirements
- #### Add the `VBA_Challenge.vbs` starter code to the Microsoft Visual Basic editor.
- #### Use the steps Refactor VBA code and measure performance to add code where indicated by the numbered comments in the starter code file.
- #### Create a resource folder to hold the pop-up messages showing the elapsed run time for the script.

# RESULTS
## Code Examples, Stock Performance Comparison and Timestamp procedure
### Code  Examples

1. #### The `tickerIndex` is set equal to zero before looping over the rows. Creating the `tickerIndex` will help access the correct index across the different arrays on VBA code.

![Capture 1](https://user-images.githubusercontent.com/76136277/104051989-06360280-51b7-11eb-9b89-803a8afef6e2.PNG)

2.  #### Create the three output arrays - `tickerVolumes` `tickerStartingPrices` `tickerEndingPrices`. The `tickerVolumes` should be a ***Long*** data type while the `tickerStartingPrices` `tickerEndingPrices` should be in ***Single*** data type**.

![Capture 2](https://user-images.githubusercontent.com/76136277/104052021-10f09780-51b7-11eb-806b-d785c2c24731.PNG)

3. #### Created a for loop to initialize the `tickerVolumes`  to ***zero***. If the next row’s ticker doesn’t match, increase the `tickerIndex`.

![Capture 3](https://user-images.githubusercontent.com/76136277/104052039-177f0f00-51b7-11eb-996b-db1e952ebb4c.PNG)

4. #### Created a loop that will loop over all the rows in the spreadsheet. In this loop, we created a script that increases the current `tickerVolumes` variable and adds the ticker volume for the current stock ticker.

![Capture 4](https://user-images.githubusercontent.com/76136277/104052062-22d23a80-51b7-11eb-9929-ec7d7f933b94.PNG)

5. #### Cell formatting - We looped through the arrays too output the `Ticker` , `Total Daily Volume`, and `Return`. Finally, stocks with positive returns are highlighted in green and negative returns in red. This is to easily identify which stocks  did well and which ones isn't do well. Also cells `C4:C15` values were changed to %.

![Capture 5](https://user-images.githubusercontent.com/76136277/104053366-17800e80-51b9-11eb-98ff-11ab651fcac3.PNG)

###  Stock Performance Comparison (2017 vs 2018)
#### In the screenshots below, we can see that 2017 has one negative return while 2018 has two negative returns. This means that, stock performance in 2017 is better than 2018.

![VBA_Challenge_2017 (2)](https://user-images.githubusercontent.com/76136277/104059196-e86e9a80-51c2-11eb-9778-45234e6ba1b8.png) ![VBA_Challenge_2018(2)](https://user-images.githubusercontent.com/76136277/104059213-ef95a880-51c2-11eb-8e24-f62e2b9dd3e2.png)

### Execution Times 

![VBA_Challenge_2017 (1)](https://user-images.githubusercontent.com/76136277/104060506-fe7d5a80-51c4-11eb-9056-e8efb242bfbc.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/76136277/104060515-00dfb480-51c5-11eb-8953-1b86a8922bd1.png)


# SUMMARY
## Advantages and Disadvantages of refactoring code
1. Advantages
- #### Refactoring makes the code more efficient, by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
- #### Refactoring helps to find bugs
- #### Refactoring improves performance  (the refactored script runs faster compared to the original script).

2. Disadvantages
- #### Refactoring is risky when the existing code doesn't have proper test cases.
- #### The cost of refactoring is higher than rewriting the code from scratch

## How do these pros and cons apply to refactoring the original VBA script?
####  Refactoring the original VBA script made the script easier to understand and user friendly. It removed code complexity ,and made the script run faster without changing its external behavior. On the other hand, if the code is not well understood, knowledge might be lost, which may require additional effort to regain this knowledge.
