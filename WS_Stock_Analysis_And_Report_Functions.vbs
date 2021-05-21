
'The main Subroutine(macro) that calls functions Format_Date_And_Sort and Generate_Summary

Sub Stock_Report_Main()
Call Format_Date_And_Sort
Call Generate_Summary
End Sub


'Format_Date_And_Sort : This functions formats the <date> field to 'date MM/DD/YYYY' and sorts the worskeet on <ticker> and <date>
Function Format_Date_And_Sort()

    For Each Current In Worksheets
    nme = Current.Name
    Worksheets(nme).Activate

    'Get Last row and column to know the range of worksheet

    Last_Row_Column_Address = Range("A1").SpecialCells(xlCellTypeLastCell).Address
    Last_Row_Column_Address_Arr = Split(Last_Row_Column_Address, "$")
    ColumnLetter = Last_Row_Column_Address_Arr(1)
    ColumnNumber = Current.Range(ColumnLetter & 1).Column
    CurrentLastRow = Last_Row_Column_Address_Arr(2)


    'this is to make sure the date is formatted for each worksheet even if an error occured in the current worksheet .
    'the error usually occurs when ithe date has already been formatted for the current worksheet
    On Error Resume Next
    For Each x In Current.Range("B2:B" & CurrentLastRow)
        x.Value = DateSerial(Left(x.Value, 4), Mid(x.Value, 5, 2), Right(x.Value, 2))
        x.NumberFormat = "mm/dd/yyyy"
        'if something goes wrong, raise an error, then
        If Err.Number <> 0 Then
            Exit For
        End If
    Next


    'Sort
    Worksheets(Current.Name).Sort.SortFields.Clear
    Current.Range("A1:" & ColumnLetter & CurrentLastRow).Sort Key1:=Range("A1"), order1:=xlAscending, Key2:=Range("B1"), order2:=xlAscending, Header:=xlYes

        
    Next

End Function


'Generate_Summary : This function reads the worksheet and generates the summary table
Function Generate_Summary()

 'Define variables
  Dim Ticker As String
  Dim Begining_Of_Year As Date ' We only need this to if we decide to display date. The worsheet has been sorted with previous subroutines'
  Dim Stock_Open_Price As Double
  Dim End_Of_Year As Date ' We only need this to if we decide to display date. The worsheet has been sorted with previous subroutines'
  Dim Stock_Close_Price As Double
  Dim Yearly_Change As Double
  
  Dim Percent_Change As Double
  Dim Total_Stock_Value As Double
  Total_Stock_Value = 0
  Begin_Date = 0
  
  Dim Summary_Table_Row As Integer
  
  Dim Current As Worksheet
  
  For Each Current In Worksheets
    nme = Current.Name
    Worksheets(nme).Activate

    ' Add Header to the summary table
      Current.Range("J1").Value = "Ticker"
      Current.Range("K1").Value = "Yearly Change"
      Current.Range("L1").Value = "Percent Change"
      Current.Range("M1").Value = "Total Stock Value"
     
      Current.Range("P2").Value = "Greatest % Increase"
      Current.Range("P3").Value = "Greatest % Decrease"
      Current.Range("P4").Value = "Greatest Total Volume"
      Current.Range("Q1").Value = "Ticker"
      Current.Range("R1").Value = "Value"
      
    Summary_Table_Row = 2
    
    
    Last_Row_Column_Address = Range("A1").SpecialCells(xlCellTypeLastCell).Address
    Last_Row_Column_Address_Arr = Split(Last_Row_Column_Address, "$")
    ColumnLetter = Last_Row_Column_Address_Arr(1)
    ColumnNumber = Current.Range(ColumnLetter & 1).Column
    CurrentLastRow = Last_Row_Column_Address_Arr(2)
    
      
    'All variables are set to empty before entering the loop
      Ticker = ""
    'Begining_Of_Year = "" 'we are not using this but we can use thes to display the date of opening price
      Stock_Open_Price = 0
    'End_Of_Year = ""      'we are not using this but we can use thes to display the date of closing price
      Stock_Close_Price = 0
      Yearly_Change = 0
      Percent_Change = 0
      Total_Stock_Value = 0
      Total_Stock_Value = 0
      Begin_Date = 0

      
      
        
  ' Loop through all credit card purchases
  For i = 2 To CurrentLastRow

    ' Check if we are still within the same ticker or not if not do the following...
    If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then

      ' Perform calculations on the ticker group
      Ticker = Current.Cells(i, 1).Value
      'End_Of_Year = Current.Cells(i, 2).Value
      Stock_Close_Price = Current.Cells(i, 6).Value
      
      Total_Stock_Value = Total_Stock_Value + Current.Cells(i, 7).Value
      Yearly_Change = Stock_Close_Price - Stock_Open_Price
      If Stock_Open_Price > 0 Then
      Percent_Change = Round(((Yearly_Change / Stock_Open_Price) * 100), 2)
      Else
      Percent_Change = 0
      End If
      
      
      
      Current.Range("J" & Summary_Table_Row).Value = Ticker
      Current.Range("K" & Summary_Table_Row).Value = Yearly_Change
      ' Conditional formatting is applied to the yearly change and/ or percent change columns
      If Yearly_Change >= 0 Then
        Current.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      ElseIf Yearly_Change < 0 Then
        Current.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      Current.Range("L" & Summary_Table_Row).Value = Percent_Change & "%"
      Current.Range("M" & Summary_Table_Row).Value = Total_Stock_Value
      
      
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
         
     
        ' Reset the variables
        Total_Stock_Value = 0
        Ticker = ""
        'Begining_Of_Year = ""
        Stock_Open_Price = 0
        'End_Of_Year = ""
        Stock_Close_Price = 0
        Yearly_Change = 0
      
        Percent_Change = 0
        Total_Stock_Value = 0
        Begin_Date = 0
    
        ' If the cell immediately following a row is the same ticker...
        Else
        'capture Begining_Of_Year, Stock_Open_Price Begin_Date is a counter .
          If Begin_Date = 0 Then
            Begining_Of_Year = Current.Cells(i, 2).Value
            Stock_Open_Price = Current.Cells(i, 3).Value
            Begin_Date = 1
          End If
          
          ' Add to the Brand Total
          Total_Stock_Value = Total_Stock_Value + Current.Cells(i, 7).Value
    
        End If
    
      Next i
      
  
    'BONUS Display : "Greatest % increase", "Greatest % decrease" and "Greatest total volume" '
    '------------------------------------------------------------------------------------------
    
    CurrentLastRowSUMM = Current.Cells(Rows.Count, "J").End(xlUp).Row
   
   
    'For i = 1 To 3
    'If i = 1 Then
    Worksheets(Current.Name).Sort.SortFields.Clear
    Current.Range("J1:M" & CurrentLastRowSUMM).Sort Key1:=Range("L1"), order1:=xlDescending, Header:=xlYes
    Current.Range("Q2").Value = Current.Range("J2").Value
    Current.Range("R2").Value = (Current.Range("L2").Value * 100) & "%"
    'ElseIf i = 2 Then
    Worksheets(Current.Name).Sort.SortFields.Clear
    Current.Range("J1:M" & CurrentLastRowSUMM).Sort Key1:=Range("L1"), order1:=xlAscending, Header:=xlYes
    Current.Range("Q3").Value = Current.Range("J2").Value
    Current.Range("R3").Value = (Current.Range("L2").Value * 100) & "%"
   ' ElseIf i = 3 Then
    Worksheets(Current.Name).Sort.SortFields.Clear
    Current.Range("J1:M" & CurrentLastRowSUMM).Sort Key1:=Range("M1"), order1:=xlDescending, Header:=xlYes
    Current.Range("Q4").Value = Current.Range("J2").Value
    Current.Range("R4").Value = Current.Range("M2").Value
    'End If
   ' Next
   
    Worksheets(Current.Name).Sort.SortFields.Clear
    Current.Range("J1:M" & CurrentLastRowSUMM).Sort Key1:=Range("J1"), order1:=xlAscending, Header:=xlYes
    
    
Next
End Function










