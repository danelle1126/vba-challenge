Attribute VB_Name = "Module1"
Sub stocks()

  ' Set up worksheet loop
  Dim WS As Worksheet
  For Each WS In Worksheets
  
  ' Set an initial variable for holding the ticker symbol
  Dim Ticker As String
  
  ' Keep track of the summary table row
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Track first row of new stock
  Dim NewStock As Boolean
  NewStock = True

  ' Set initial variables for holding the open and close stock prices
  Dim Open_Price As Double
  Open_Price = 0
  Dim Close_Price As Double
  Close_Price = 0

  ' Keep track of total stock volume for year
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Unkown number of rows
  LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all the stock data
  For i = 2 To LastRow

    ' Check if we are still within the same stock, if it is not...
    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

      ' Set ticker symbol
      Ticker = WS.Cells(i, 1).Value

      ' Add to the Stock Volume Total
      Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value

      ' Print the ticker symbol in the Summary Table
      WS.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the stock volume to the Summary Table
      WS.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      
      ' Capture close price and calculate yearly change and percent change
      Close_Price = WS.Cells(i, 6).Value
      WS.Range("J" & Summary_Table_Row).Value = Close_Price - Open_Price
      If Open_Price > 0 Then
      WS.Range("K" & Summary_Table_Row).Value = (Close_Price - Open_Price) / Open_Price
      Else
      WS.Range("K" & Summary_Table_Row).Value = 0
      End If
    
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
      ' Reset total stock volume and NewStock
      Total_Stock_Volume = 0
      NewStock = True
 
    ' If the cell immediately following a row is the same stock...
    Else
      
      ' Add current volume to the total stock volume
      Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
      
      ' Capture open price for new stock
      If NewStock = True Then
           Open_Price = WS.Cells(i, 3).Value
      End If
      
      NewStock = False
      
    End If

  Next i
  
  Next WS

End Sub
