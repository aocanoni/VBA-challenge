Within the VBAChallenge TurnIn folder are all the necessary files for the assignment

Challenge2VBAfile.bas is my VBA script
  - This script has 1 module with 3 subs, but the last two subs are called at the end of the first sub
  - The first sub (multiple_year_stock_data) functions to input the first summary row table within the excel file (so, ticker, yearly change, percentage change, and total stock volume) and it lastly calls the other two subs.
  - The second sub (yearly_change_color) formats the first sub's yearly change column to include colored boxes according to the inputted values. 
  - The last sub (ticker_greatest) inputs a small table including the greatest % increase, decrease, and total volume out of all of the summary table. Lastly, it formats all of the worksheets so that all input data autofits to the column width.
  - Each sub iterates through each worksheet.
  - Code Source: the majority of the code is my work based on the provided unsolved and solved activities given to us by the bootcamp. The class instructor helped with pointing out redundant code and other questions I had (such as percentage and currency formatting, to which they showed me the macro recording function), and lastly I referenced the Microsoft website in order to find out how to use the autofit columns function. 

There are three screenshots of the excel workbook as well
  - Each screenshot is named according to their worksheet name (2018, 2019, and 2020) and should appropriately display the top of each worksheet which should include all displayed values, column names, and the top of each summary table.
