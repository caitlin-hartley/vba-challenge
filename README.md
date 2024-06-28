# vba-challenge
### hw2, stock analysis
### VBA script that analyzes quarterly stock data
- code folder contains vbs code and xlsm sheet

[Values](https://github.com/caitlin-hartley/vba-challenge/blob/main/README.md#finding-and-calulating-values-in-each-sheet)
  
## VBA script begins by going through each sheet:
  - Creates the summary table headers
  - Calculates the min and max date for each quarter
  - Finds the last row of the sheet
  - Starts the summary row counter
![sheets](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/sheet_loop.png)

## Finding and calulating values in each sheet:
- Find unique ticker values
- Add up stocker volume for that unique ticker
- Finds opening value based on minimum date for each ticker
- Finds closing value for each ticker
- Calculates the quarterly change in price for each ticker
- Calculates percent change in stock price over the quarter (change in price / opening price)
![values](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/values_loop.png)

## Adding values to summary table:
- Add values to each summary table
- Format percentages
- Color codes the quarterly change based on whether the increase was positive, negative, or zero
- Increases summary row counter by 1 to move to the next row of the summary table
![summary](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/summary_table_loop.png)

## Goes through summary table to find:
- the stock with the greatest percent increase
- the stock with the greatest percent decrease
- the stock with the greatest total volume
- the name of the stock with these values
- Add to smaller stats table
![greatest](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/greatest_stock_loop.png)

## The results for the different quarters are below: 

Q1:
![Q1](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/q1_stock_results.png)

Q2:
![Q2](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/q2_stock_results.png)

Q3:
![Q3](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/q3_stock_results.png)

Q4:
![Q4](https://github.com/caitlin-hartley/vba-challenge/blob/main/images/q4_stock_results.png)
