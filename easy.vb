Sub Test()

' Assigning a variable for the stock names
Dim ticker As String
Cells(1, 10).Value = "Ticker"

' Assigning a variable for the total volume of stock
Dim total As Double
Cells(1, 11).Value = "Total Stock Volume"

' Keeping track of each stock symbol
Dim summarytablerow As Integer
summarytablerow = 2
 
'Finding the last non-blank cell in column A(1)
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Looping through all stock volumes for each stock symbol
For i = 2 To lastrow
 
  ' Checking if we are still within the same stock symbol
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Setting the stock symbol name
    ticker = Cells(i, 1).Value
    
    ' Summing up to total volume
    total = total + Cells(i, 7).Value
    
    ' Printing the stock symbol 
    Range("J" & summarytablerow).Value = ticker
    
    ' Printing the total volume 
    Range("K" & summarytablerow).Value = total
     
    ' Adding one to the summary table row
    summarytablerow = summarytablerow + 1
     
    ' Resetting the total volume for each stock symbol
    total = 0

  'If we are not within the same stock symbol then
  Else

    ' Summing up the total volume
    total = total + Cells(i, 7).Value

  'Closing if statement
  End If

'Ending loop
 Next i

End Sub


