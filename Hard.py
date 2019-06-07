Sub Hard()


    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
 For Each ws In Worksheets
 
     ' --------------------------------------------
     ' return the total volume each stock had over that year + display the ticker symbol to coincide with the total stock volume
     ' --------------------------------------------
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  
  'Create variables to hold ticker, total volume and file name
  Dim Ticker As String
  Dim Stock_Volume As Double
  Dim WorksheetName As String
  Dim Open_Value As Double
  Dim Close_Value As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
    
  
  'Initialize the variables
  Ticker = 2
  Stock_Volume = 0
  Open_Value = 0
  Close_Value = 0
  Percent_Change = 0
  

  ' Loop through rows in the column
  LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
  
  ' Look for when the tickler change in column A
  
  For i = 2 To LastRow
  
     Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
     Open_Value = Open_Value + ws.Cells(i, 3).Value
     Close_Value = Close_Value + ws.Cells(i, 6).Value
     Yearly_Change = Close_Value - Open_Value
     'Percent_Change = (Yearly_Change \ Open_Value) * 100
     'Percent_Change = (ws.Cells(i, 3).Value / ws.Cells(i, 6).Value)
     
    ' Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
       ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
       ws.Cells(Ticker, 13).Value = Open_Value
       'ws.Cells(Ticker, 14).Value = Close_Value
       ws.Cells(Ticker, 10).Value = Yearly_Change
       ws.Cells(Ticker, 11).Value = Percent_Change
       ws.Cells(Ticker, 12).Value = Stock_Volume
       
       Open_Value = 0
       Close_Value = 0
       Stock_Volume = 0
       Yearly_Change = 0
       'Percent_Change = 0

    Ticker = Ticker + 1
    
    End If

  Next i
 
 ' Conditional formatting that will highlight positive change in green and negative change in red.
   
  For i = 2 To Range("J" & Rows.count).End(xlUp).Row
    If ws.Cells(i, "J") > 0 Then
        ws.Cells(i, "J").Interior.ColorIndex = 4
    Else
        ws.Cells(i, "J").Interior.ColorIndex = 3
    End If
' End of coloring formatting

Next
  
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
   
  For i = 2 To Range("K" & Rows.count).End(xlUp).Row
 
        ws.Cells(i, "K").Value = (Cells(i, 10).Value / Cells(i, 13).Value) * 100
        ws.Columns("K:K").NumberFormat = "0.00%"
                
   
  Next
' End of percent change and formatting

Dim rng As Range
Dim Max As Double
Dim Min As Double
Dim Volumen As Variant



'Set range from which to determine largest value
'Set rng = Sheet1.Range("K1:K1000")

'Worksheet function MAX returns the largest value in a range

fin = Cells(Rows.count, 2).End(xlUp).Row

Max = Application.WorksheetFunction.Max(Range("K2:K" & fin))
Min = Application.WorksheetFunction.Min(Range("K2:K" & fin))
Volumen = Application.WorksheetFunction.Max(Range("L2:L" & fin))

'Displays largest value
'MsgBox Max
'MsgBox Min

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


ws.Cells(2, 17).Value = Max
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).Value = Min
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 17).Value = Volumen

  Next ws
  
  

End Sub
