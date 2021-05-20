

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
The excel raw data files are also loacted in the folder.  
Use the sheet alphabetical_testing.xlsx while testing your code. This data set is smaller and will allow faster testing. 
The macro on this file excecutes in undweer a minute .


