
#  VBA scripting to analyze stock market data
#### Project Description

* VBA Excel macro that reads stocks from the workboox with each worksheet containing one year's data and output the following information:
  * The ticker symbol.
  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The total stock volume.
  
* Conditional formatting that highlights positive change in green and negative change in red.
* It also creates an additional summary table with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

![Stock_Report_2014_C](https://user-images.githubusercontent.com/81383838/119061370-d6371100-b999-11eb-8df3-9e005be59635.jpg)

#### Sequence of execution
##### The main Subroutine(macro) calls two functions Format_Date_And_Sort and Generate_Summary

Sub Stock_Report_Main()\
Call Format_Date_And_Sort\
Call Generate_Summary\
End Sub


###### Format_Date_And_Sort():
This function formats the 'date' field to 'MM/DD/YYYY' and sorts the worksheet on 'ticker' and 'date' fields. 
###### Generate_Summary() :
This function loops through the the worksheet, reads data from each row and generates summary table from the data.
 
#### Execution:
  * The code is in the file "WS_Stock_Analysis_And_Report_Functions.vbs" which can be located in files section.
  * The excel raw data files are also loacted in the same folder.
     * I tested the macro on alphabetical_testing.xlsm.
     * The macro on this file executed in under ONE minute.
     * You can try this by downloading the raw data file here and creating a module using the code from WS_Stock_Analysis_And_Report_Functions.vbs.
     * Please copy all the lines of  the code that includes the main subroutine and the two functions that it calls.

 * I have also tested the same macro on the much larger data file, and it executed in just under 6 mins. 
    * Most of the execution time is taken for date formatting. Once that has been formatted, the processing to generate the report is less than 2 minutes.
    * You can try this by downloading the data file here and using the code from WS_Stock_Analysis_And_Report_Functions.vbs. 
    * to make the excution faster, i have already run the macro to format date and sort the worksheet. 
    * Please copy all the lines of the code that includes the main subroutine and the two functions that it calls.
 
 * If you would like to try it on a raw file for fun, here is the data file. 
    * The execution time is up to 6 mins.
    * Follow the same instructions as above and copy all of the code from WS_Stock_Analysis_And_Report_Functions.vbs.
    * Be patient while the file is being processed. Remember "Patience is Golden !"
 
###### eRRORS !!!! 
* REST ASSURED, THEY HAVE BEEN HANDLED. Just Follqw these instructions ....
 
###### Once you have added the code to excel workbooks as a module or just as a macro for the work book, you will able to run the macro Stock_Report_Main().

#### Important: Add the vbs code to the workbook or to the module not individual worksheets.
