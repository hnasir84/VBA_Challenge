Attribute VB_Name = "Module1"
' Create a script that loops through all the stocks for one year and outputs the following information:

    ' The ticker symbol

    ' Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

    ' The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

    ' The total stock volume of the stock.
    
    
Sub StockMarket():

'Set  up a variable  that is the type of worksheet
 Dim Sheet As Worksheet
 
 
'create a loop that goes through all the sheet int he workbook
 For Each Sheet In ThisWorkbook.Worksheets
 

' assign columns I, J, K and L to hold the summary table
Sheet.Range("I1").Value = "Ticker"
Sheet.Range("J1").Value = "Yearly Change"
Sheet.Range("K1").Value = "Percnt Change"
Sheet.Range("L1").Value = "Total Stock volume"


' set Variable called Percentage_Change
 Dim Percentage_Change As Double


' set initial variables for holding Ticker Name in the summary tble
  Dim Ticker_Name As String
  
'Set initial variable to hold Yearly Change
 Dim Yearly_change As Double


'Set initial variable to hold the opening price
 Dim Opening_Price As Double

'Set initial variable to hold the Closing price
 Dim Closing_Price As Double
 

'set initial variable to hold the total volume per ticker
 Dim Total_Stock_Volume As Double
 Total_Stock_Volume = 0
 
' set a variable to keep track of the location for each Ticker symbol in the summary table
 Dim Summary_table_row As Integer
 Summary_table_row = 2


'find the last row in the main table, Column A
 LastRow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row


'Find the last row in the summary table
 LastRowSummary = Sheet.Cells(Rows.Count, "K").End(xlUp).Row




'set the opening price value
 Opening_Price = Sheet.Cells(2, 3).Value



' Loop through all rows in the table to create the summary table
 For i = 2 To LastRow


' Check if we are still within the same Ticker Name, if not..

  If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value Then
  
  
    ' Set the Ticker name
      Ticker_Name = Sheet.Cells(i, 1).Value
    
    
    ' set the closing price
      Closing_Price = Sheet.Cells(i, 6).Value
   
    ' Claculate the yearly change
      Yearly_change = Closing_Price - Opening_Price
    
 
   
    'Print the yearly Change
     Sheet.Cells(Summary_table_row, 10).Value = Yearly_change
    
    
    'Calculate the percentage between closing and opening prices
     Percentage_Change = Yearly_change / Opening_Price
   
   
   
    'Print the Percentage change in Summary Table and round the outcome
     Sheet.Cells(Summary_table_row, 11).Value = Round(Percentage_Change, 4)
  
   
   
    'Format the percentage_Change output in the summary table to percentage
     Sheet.Cells(Summary_table_row, 11).NumberFormat = "0.00%"
    

    'Update the opening price for the next Ticker symbole in the summary table
     Opening_Price = Sheet.Cells(i + 1, 3).Value
  
 

    ' Add to the volume total
      Total_Stock_Volume = Total_Stock_Volume + Sheet.Cells(i, 7).Value
  
  
    ' Print Ticker symbole in the summary table
      Sheet.Cells(Summary_table_row, 9).Value = Ticker_Name
  
    ' Print the total stock volume of the ticker to the summary table
      Sheet.Cells(Summary_table_row, 12).Value = Total_Stock_Volume
   
    'Add 1 to the summary table row
     Summary_table_row = Summary_table_row + 1
  
    ' Reset the Total_Stock_Volume and Yearly change to 0 in order to count the next ticker
      Total_Stock_Volume = 0
      Yearly_change = 0
   


' if the cells immediatly following a  row hold the same ticker then ...
  
  
  Else
  
  
  ' add to the Total_Stock_Volume of the same ticker
    Total_Stock_Volume = Total_Stock_Volume + Sheet.Cells(i, 7).Value

  
   
 End If
 

  
Next i

 


' Apply conditional formatting that will highlight positive change in green and negative change in red

' Create loop  that will run from the second row to the last row in the summary table


For L = 2 To LastRowSummary

' if the yearly change is positive, the cell interior will be green
  If Sheet.Cells(L, 11).Value > 0 Then


  Sheet.Cells(L, 11).Interior.ColorIndex = 4


  Else

' if the yearly change is Negative, the cell interior will be Red
  Sheet.Cells(L, 11).Interior.ColorIndex = 3


End If



Next L



'----------------------------------------------------------------------------
                      'BOUNS
                      
' Add functionality to your script to return the stock with the
' "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

'-----------------------------------------------------------------------------



'Set the table that will hold the summarized data

Sheet.Range("P1").Value = "Ticker"
Sheet.Range("q1").Value = " Value"
Sheet.Range("O2").Value = "Greatest  % Increase"
Sheet.Range("O3").Value = "Greatest  % Decrease"
Sheet.Range("O4").Value = "Greates Total Volume"


'  set variables to hold the of  greatest increase, greatest Decrease
'  Set variables to hold the ranges of interest
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Tickerindex As Integer
Dim PercentChangeRange As Range
Dim GreatestTotalVolRange As Range



' Set the range where the max and min percent change will be found
Set PercentChangeRange = Sheet.Range("k2", "K" & LastRowSummary)

' Find the greatest increase using the Max function
  Greatest_Increase = WorksheetFunction.Max(PercentChangeRange)

' Use the Match function to find the cell that hold the greatest increase
 Tickerindex = WorksheetFunction.Match(Greatest_Increase, PercentChangeRange, 0)

' Print the greatest increase & corresponding ticker symbol in the summary table
 Sheet.Range("Q2").Value = Greatest_Increase
 
 Sheet.Range("p2").Value = Sheet.Range("I" & Tickerindex + 1).Value
 
' Find the greatest increase using the Min function
  Greatest_Decrease = WorksheetFunction.Min(PercentChangeRange)


' Use the Match function to find the cell that hold the greatest Decrease
  Tickerindex = WorksheetFunction.Match(Greatest_Decrease, PercentChangeRange, 0)

' Print the greatest Decrease & corresponding ticker symbol in the summary tabl
  Sheet.Range("Q3").Value = Greatest_Decrease
 
  Sheet.Range("p3").Value = Sheet.Range("I" & Tickerindex + 1).Value
 
 
 
' Set the range where the greatest total volume will be found
  Set GreatestTotalVolRange = Sheet.Range("L2", "L" & LastRowSummary)
 
' Find the greatest total volume increase using the Max function
  Greatest_Total_Volume = WorksheetFunction.Max(GreatestTotalVolRange)
 
' Use the Match function to find the cell that hold the greatest total volume
  Tickerindex = WorksheetFunction.Match(Greatest_Total_Volume, GreatestTotalVolRange, 0)
 
 
' Print the greatest Decrease & corresponding ticker symbol in the summary tabl
  Sheet.Range("Q4").Value = Greatest_Total_Volume
 
  Sheet.Range("p4").Value = Sheet.Range("I" & Tickerindex + 1).Value

 ''''''''''''''''''''''''''''''''''''''''''''''''
 ' Apply formatting to the tables
 ''''''''''''''''''''''''''''''''''''''''''''''''
 
 ' Apply Autofit to all used Columns
 
  Sheet.UsedRange.EntireColumn.AutoFit
 
  'Format the output summary of Greatest_Increase &  to percentage
   Sheet.Range("Q2:Q3").NumberFormat = "0.00%"

 
Next Sheet


End Sub
