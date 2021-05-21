
#  VBA scripting to analyze stock market data
#### Project Description

* VBA Excel macro that reads stocks from the workboox with each worksheet containing one year's data and output the following information:
  * The ticker symbol.
  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The total stock volume.
  
* Conditional formatting that highlights positive change in green and negative change in red.
* It also creates an additional summary table with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

![Stock_Report_14](https://user-images.githubusercontent.com/81383838/119068685-4d74a100-b9aa-11eb-8423-27c711b66c27.jpg)
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
  * The code is in the file ![WS_Stock_Analysis_And_Report_Functions.vbs](https://github.com/V-MalM/VBA-challenge/blob/main/WS_Stock_Analysis_And_Report_Functions.vbs) which can be located in files section.
  * Add the code to excel workbook by creating a module or by selecting 'Thisworkbook'.
  * Save and you will able to run the macro "Stock_Report_Main()".
     * I tested the macro on ![alphabetical_testing.xlsm](https://github.com/V-MalM/VBA-challenge/blob/main/alphabetical_testing.xlsm).
     * It took less than 1 minute to execute and show the results.

 * I have also tested the same macro on the "Multiple_year_stock_data.xlsm", and it executed in 8 mins. 
    * Most of the execution time is taken for date formatting. Once that has been formatted, the processing to generate the report is less than 2 minutes.
    * I had to Remember "Patience is Golden !"

#### Results
#### 2014
![Stock_Report2014](https://user-images.githubusercontent.com/81383838/119071689-e22dcd80-b9af-11eb-8182-6806b0fbdb2f.jpg)
#### 2015
![Stock_Report2015](https://user-images.githubusercontent.com/81383838/119071694-e5c15480-b9af-11eb-96fd-7c96c1512f5a.jpg)
#### 2016
![Stock_Report2016](https://user-images.githubusercontent.com/81383838/119071702-ea860880-b9af-11eb-90f1-34dea540b087.jpg)


###### Once you have added the code to excel workbooks as a module or just as a macro for the work book, you will able to run the macro Stock_Report_Main().
#### Important: Add the vbs code to the workbook or to the module not individual worksheets.
###### eRRORS !!!! What ERRORS ????
* REST ASSURED, THEY HAVE BEEN HANDLED. Just Follow these detailed instructions ....
 
