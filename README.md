# VBA-challenge
Using VBA scripting to analyze real stock market data


### Background
For this project, I created a VBA (Visual Basic) script to analyze some stock market data. The data is inside a Microsoft Excel workbook and includes stock data for three years (2014, 2015, and 2016). 

Due to the size of the Excel workbook, I was not able to upload that file to this GitHub repository.

Regardless, the multiple year stock data along with the test data used to create this script can be found here:

https://rice.bootcampcontent.com/Rice-Coding-Bootcamp/RU-HOU-DATA-PT-12-2019-U-C/tree/master/Homework/02-VBA-Scripting/Instructions/Resources

### Testing
I ran this script on both the testing Excel workbook (alphabetical_testing.xlsx) and on the final multiple year stock workbook (multiple_year_stock_data.xlsx).


### About the Script
To run the script, you can do the following:
- Click the Developer tab.
- Click Visual Basic to open the Visual Basic editor.
- Inside the Visual Basic editor, click File > Import File and import the CalculateStockStats.bas file in this repository.


As the script runs, it loops through all the stocks for one year for each run and takes the following information:
- The ticker symbol
- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.
- It applies conditional formatting by highlighting positive yearly change values in green and negative yearly change values in red.
- It returns the stock with the greatest percent increase, greatest percent decrease, and greatest total volume

### Sample Output
After the script has completed, go to the Excel workbook, and you should see the results of the script.
![solution] (https://rice.bootcampcontent.com/Rice-Coding-Bootcamp/RU-HOU-DATA-PT-12-2019-U-C/raw/master/Homework/02-VBA-Scripting/Instructions/Images/moderate_solution.png)

![solution_hard] (https://rice.bootcampcontent.com/Rice-Coding-Bootcamp/RU-HOU-DATA-PT-12-2019-U-C/raw/master/Homework/02-VBA-Scripting/Instructions/Images/hard_solution.png)


