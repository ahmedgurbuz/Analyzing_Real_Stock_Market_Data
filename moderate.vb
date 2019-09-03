Sub Test()

' Assigning a variable for the stock names
Dim ticker As String
Cells(1, 10).Value = "Ticker"

' Assigning a variable for the total volume of stock
Dim total As Double
Cells(1, 13).Value = "Total Stock Volume"

'Assigning Yearly Change and Percent Change
Dim yearlychange As Double
Dim percentchange As Double
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"

' Assigning a variable to keep track of each stock symbol
Dim summarytablerow As Integer
summarytablerow = 2
 
'Finding the last non-blank cell in column A
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Assigning the first price for the first stock
Dim firstprice As Double
firstprice = Cells(2, 3).Value

' Looping through all stock volumes for each stock symbol
For i = 2 To lastrow

    ' Checking if we are still within the same stock symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
     
        ' Assigning the stock names
        ticker = Cells(i, 1).Value
     
        ' Summing up to the total volume
        total = total + Cells(i, 7).Value
     
        ' Printing the stock names
        Range("J" & summarytablerow).Value = ticker
     
        ' Printing the total volume
        Range("M" & summarytablerow).Value = total
     
        ' Assigning the last price for each stock
        Dim lastprice As Double
        lastprice = Cells(i, 6).Value

        ' Printing the Yearly Change
        yearlychange = (lastprice - firstprice)
        Range("K" & summarytablerow).Value = yearlychange

            'Preventing divided by 0 when the firstprice is 0
            If firstprice = 0 Then
                percentchange = 0
                
            'Calculating the Percentage
            Else
                percentchange = (yearlychange / firstprice)
            
            'Closing if statement
            End If    


        'Printing the percentage change
        Range("L" & summarytablerow).Value = percentchange
        Range("L" & summarytablerow).NumberFormat = "0.00%"

        ' Counter for summary table
        summarytablerow = summarytablerow + 1
     
        ' Resetting the Total Volume for each ticker name
        total = 0
        
        ' Assigning the first price for next stock
        firstprice = Cells(i + 1, 3).Value

    'If we are not within the same stock symbol then
    Else
        ' Summing up the total volume
        total = total + Cells(i, 7).Value

    'Ending if statement
    End If

'Ending loop
Next i
 
'Finding the last non-blank cell in column K
Dim lastrowK As Long
lastrowK = Cells(Rows.Count, 11).End(xlUp).Row
    
'Using loop / filling cells with colors depending on the value
For i = 2 To lastrowK

    'if cell value is greater than zero fill with green color
    If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
    
    'if cell value is not greater than zero fill with red color
        Cells(i, 11).Interior.ColorIndex = 3
    
    'Closing if statement
    End If

'Closing loop
Next i

End Sub
