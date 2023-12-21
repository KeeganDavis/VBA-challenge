# VBA-challenge

## Description
This VBA script analyzes stock tickers over the course of a year and calculates the change in dollar amount, percent change of the stock price, and the total stock volume. The script then finds the ticker with the greatest percent increase, the greatest percent decrease, and the greatest total volume. All negative values in the yearly change and percent change columns are filled in red and all positive values in the yearly change and percent change columns are filled in green. The script will run for all sheets containing daily ticker data for the year.

## Installation
Clone the repository: `git clone git@github.com:KeeganDavis/VBA-challenge.git`

## Usage
To use this code, run the AllSheetsAnalysis() Sub and the data will be analyzed for each ticker on every sheet. The column labels, row labels, and data will be added to the cells. Conditional formatting will be applied so that the negative values in the yearly change and percent change columns are filled in red and all positive values in the yearly change and percent change columns are filled in green. To clear all of the newly added data, run the ResetAll() Sub and the only data remaining will be the original yearly ticker data.

## Code Source
#### Within Stock_script
-line 11 and 47: How to find the last row containing data (https://www.wallstreetmojo.com/vba-last-row/)
-line 96-98 and 118-120: How to apply code to all worksheets (https://excelchamps.com/vba/loop-sheets/)
-basically all lines of code with ws: How to pass worksheet as a parameter to apply to all worksheets (https://stackoverflow.com/questions/31706678/pass-worksheet-as-parameter-and-then-utilize-that-parameter-as-a-variable)
-line 90: How to automaticall adjust column size (https://stackoverflow.com/questions/24154232/vba-to-select-all-columns-in-a-worksheet-and-auto-adjust-all-columns-width-in-ex)
-line 91 and 102-105: How to format cells to percent, number, and currency (https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba)
-lines 99-101: Calling a sub with arguments (https://stackoverflow.com/questions/56259496/calling-a-sub-in-vba-with-multiple-arguments)
-lines 106-112: Conditional formatting (https://www.automateexcel.com/vba/conditional-formatting/)
-lines 122-123: Clear formatting (https://stackoverflow.com/questions/31884818/clear-contents-and-formatting-of-an-excel-cell-with-a-single-command)
-line 121: Using set (https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)