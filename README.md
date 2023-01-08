## Goal
* Search through the multiyear stocks data to obtain ticker symbols
* Sort through and list ticker symbols
* Calculate and log the difference of the opening value of each ticker at the beginning of the year and its value at year's end
* Analyze the percentage of change by ticker and log the ticker with the largest increase, decrease, and volume based on information from dataset
* Analyze this information for each year listed in the dataset


## Method
* Create For loop to ensure code is operational for each worksheet
* Set variables for all desired items and values (e.g. worksheet name, data rows, calculated values, etc.)
* Create column headers for calculated values across all worksheets
* Create for loop to search through full list of tickers
* Return ticker name to summary column each time it changes from the row above it
* Return calculated yearly change, percentage change, and total volume by ticker
* Set conditional formatting to set positive change to green and negative to red
* Obtain and log amount of increase and decrease in value and total volume by ticker
* Log and overwrite values if they exceed the previous (until the greatest increase, decrease, and total volume are found)

## Summary and Results
* In 2018, the ticker THB demonstrated the greatest percent increase in value, RKS demonstrated the greatest decrease, and QKN demonstrated the greatest total volume

![screenshot1](https://github.com/Grimmandrewj/Stocks_Analysis/blob/main/Screenshots/2018%20screenshot.jpg)

* In 2019, the ticker RYU demonstrated the greatest percent increase in value, RKS again had the greatest decrease, and ZQD had the greatest total volume

![screenshot2](https://github.com/Grimmandrewj/Stocks_Analysis/blob/main/Screenshots/2019%20screenshot.jpg)

* In 2020, YDI demonstrated the greatest percent increase in value, VNG the greatest decrease, and QKN again had the greatest total volume

![screenshot3](https://github.com/Grimmandrewj/Stocks_Analysis/blob/main/Screenshots/2020%20screenshot.jpg)

* There was no consistent ticker that demonstrated the greatest value increase over the three years of provided data.  However, RKS twice demonstrated the greatest decrease in value (2018 and 2020), suggesting it was a poorly-performing at the time of this dataset.  
* QKN demonstrated the greatest total volume in 2018 and 2020, suggesting it was the ticker with the most liquidity at the time this dataset was recorded


