
# VBA-challenge

#### VBA scripting to analyze stock market data

#### Project Description

* Created a VBA excel macro that reads each worksheet and formats the date column and also sorts in order of Ticker and Date.
* Created a VBA excel macro that reads stocks form excel workboox with each sheet containing one year's data and output the following information.
  * The ticker symbol.
  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The total stock volume.
  
* Conditional formatting has been applied that highlights positive change in green and negative change in red.
* It also creates a table with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

![Stock_Report_2014_C](https://user-images.githubusercontent.com/81383838/119061370-d6371100-b999-11eb-8df3-9e005be59635.jpg)

#### Sequence of execution
The main Subroutine(macro) calls two functions Format_Date_And_Sort and Generate_Summary

Sub Stock_Report_Main()
Call Format_Date_And_Sort
Call Generate_Summary
End Sub

Format_Date_And_Sort
This functions formats the <date> field to 'date MM/DD/YYYY' and sorts the worskeet on <ticker> and <date>
 
Generate_Summary : 
This function reads the worksheet and generates the summary table

The code is in the file WS_Stock_Analysis_And_Report_Functions.vbs which can be located in files section.
The excel raw data files are also loacted in the same folder.  
I tested the macro on the sheet alphabetical_testing.xlsx while testing the code. This data set is smaller and will allow faster testing. 
The macro on this file excecuted in under a minute.
You can try this by downloading the raw data file here and creating a module using the code from WS_Stock_Analysis_And_Report_Functions.vbs. Please copy all the lines of the code that includes the main subroutine and the two functions that it calls.
 
I have also tested the same macro on the much larger data file and it executed in just under 6 mins. Most of the execution time is taken for date formatting. Once that has been formatted, the execution to generate the report is less than 2 minutes.
You can try this by downloading the raw data file here and using the code from WS_Stock_Analysis_And_Report_Functions.vbs. to make the excution faster, i have already run the macro to format the date and sort the worksheet. Please copy all the lines of the code that includes the main subroutine and the two functions that it calls.
 
But you would like to try it on a raw file for fun ,here is the raw data file. please follow the same instructions and copy all of the code from WS_Stock_Analysis_And_Report_Functions.vbs.
 
Once you have added the code to excel workbooks as a module or just as a macro for the work book, you will able to run the macro Stock_Report_Main().

## Important: Add the vbs code to the work book or to the module not individual worksheets
