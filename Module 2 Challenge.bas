Attribute VB_Name = "Module1"
Sub Stock_Data()
    
    'Initializing process to repeat macro across all worksheets
    Dim w As Long
    'Turning off screen updating
    Application.ScreenUpdating = False
    'Initialize For loop to activate worksheets one by one
    For w = 1 To Worksheets.Count
    Sheets(w).Select

    'Turn on screen updating after activating each sheet
    Application.ScreenUpdating = True

    'Printing header titles for worksheets
    Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    Range("Q1").Value = "Value"
    Range("P1").Value = "Ticker"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    '  Set variables for worksheets
    Dim Stock_Name As String
    Dim Stock_Volume As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim lastrow As Long
    Dim ticker As String
    Dim IncTicker As String
    Dim decTicker As String
    Dim volTicker As String
    Dim greatestVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim adjacentstring As String
    Dim maxcell As Range
    Dim maxVal As Double
    
    
    'Set Ticker for Greatest Increase
    greatestIncrease = 0
    IncTicker = " "
    
    ' Set Ticker for Greatest Decrease
    
    GreatsestDecrease = 0
    decTicker = " "
    
    ' Set Ticker for Greatest Volume
    
    greatestVolume = 0
    volTicker = " "
    
       
    '  Keep track of the location for each stock in the summary table
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Stock_Volume = 0
    
    '  Using lastrow function to decipher last row of data
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Initializing for loop
    For i = 2 To lastrow
    
    'Check if we are still with the same stock, if it is not..
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
  
    ' Set the Stock name
    Stock_Name = Cells(i, 1).Value
    ' Set open price
    openprice = Cells(i, 3).Value
    ' Reset stock volume
    Stock_Volume = 0
    
    End If
    
    ' Finding the total stock volume by looping through the values in column 7
    
    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    '  Using If statement to find the closeprice and yearly change of each stock
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    closeprice = Cells(i, 6).Value
    
    yearlychange = closeprice - openprice
    
    If openprice <> 0 Then
    
    'Calculator the percent change
    percentchange = yearlychange / openprice
    
    Else
    
    percentchange = 0
    
    End If
    
    ' Using If statement to find the greatest increase in percent changed
    
    If percentchange > greatestIncrease Then
    greatestIncrease = percentchange
    IncTicker = Cells(i, 1).Value
    End If
    
    ' Using If statement to find the greatest decrease in percent changed
    
    If percentchange < greatestDecrease Then
    greatestDecrease = percentchange
    decTicker = Cells(i, 1).Value
    End If
    
    'Print the Stock in the Summary Table
    
    Range("I" & Summary_Table_Row).Value = Stock_Name
    
    'Print the Stock Volume to the Summary Table
    Range("L" & Summary_Table_Row).Value = Stock_Volume
    
    'Print the Yearly Change to the Summary Table
    
    Range("J" & Summary_Table_Row).Value = yearlychange
    Range("J" & Summary_Table_Row).Style = "Currency"
    
    
   ' Using conditional formatting to assign positive value as green and negative value as red
    If yearlychange > 0 Then
        Range("J" & Summary_Table_Row).Interior.Color = vbGreen
            Else
        Range("J" & Summary_Table_Row).Interior.Color = vbRed
            End If
    
    'Print the Percent Change to the Summary Table
    Range("K" & Summary_Table_Row).Value = percentchange
    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    ' Using conditional formatting to assign positive value as green and negative value as red
    If percentchange > 0 Then
        Range("K" & Summary_Table_Row).Interior.Color = vbGreen
            Else
        Range("K" & Summary_Table_Row).Interior.Color = vbRed
            End If
    
    Range("P2").Value = IncTicker
    Range("P3").Value = decTicker
    'Range("P4").Value = volTicker
    
    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    'Find the Max Stock Volume and printing the corresponding ticker
    Set Rng = Range("L:L")
    maxVal = WorksheetFunction.Max(Rng)
    Set maxcell = Rng.Find(What:=maxVal)
    adjacentstring = maxcell.Offset(0, -3).Value
    Range("P4") = adjacentstring
    
    'Finding the Greatest % Increase in stock price
    Set Rng = Range("K:K")
    maxvalue = Application.WorksheetFunction.Max(Rng)
    Range("Q2").Value = maxvalue
    Range("Q2").NumberFormat = "0.00%"
    
    'Finding the Greatest % Decrease in stock price
    Dim minvalue As Variant
    Set Rng = Range("K:K")
    minvalue = Application.WorksheetFunction.Min(Rng)
    Range("Q3").Value = minvalue
    Range("Q3").NumberFormat = "0.00%"
    'Finding the greatest stock volume
    Set Rng = Range("L:L")
    maxvalue = Application.WorksheetFunction.Max(Rng)
    Range("Q4").Value = maxvalue
    
    
  Else
  
  
   
    
    
  End If
   
Next i

Next w
    
    
End Sub
