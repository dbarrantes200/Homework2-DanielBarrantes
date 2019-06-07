sub Alphabetical_Testing()


    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
 For Each ws In Worksheets
 
     ' --------------------------------------------
     ' return the total volume each stock had over that year + display the ticker symbol to coincide with the total stock volume
     ' --------------------------------------------

  
  'Create variables to hold ticker, total volume and file name
  Dim Ticker As String
  Dim Stock_Volume As Double
  Dim WorksheetName As String
  
  'Initialize the variables
  Ticker = 2
  Stock_Volume = 0
  

  ' Loop through rows in the column
  LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
  
  ' Look for when the tickler change in column A
  
  For i = 2 To LastRow
  
     Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
    ' Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
       ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
       ws.Cells(Ticker, 10).Value = Stock_Volume
       Stock_Volume = 0

    Ticker = Ticker + 1
    
    End If

  Next i
  
  Next ws

End Sub
