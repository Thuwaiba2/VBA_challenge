# VBA_challenge
In this challenge i used the given data to generate yearlychange, perscntage change and total stock volume.
I used codes from class to generate the last code "lastrow=Cells(Rows.count, 1).End(xlUp).Row"
Then i wrote  a code referencing ("Interior.ColorIndex") from a code in class to fill cells with color for negative and positive value. for this i tried using "negative" and "positive" but they didnt work. I tried using less than "0" and greate than "0" but i wasnt succesful with tthat too. in module 2 but i font know why that hasnt worked
I was not able to generate a code that will work for the greatest increase and and decrease and greatest total volume. I tried to use the maximum and minimum function.

## VBA SCRIPT
Sub mutipleyearstockdata()

    

    Dim tickersymbol As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim vol As Variant
    Dim lastRow As Long
    
    Dim r As Integer

    firstrow = 2

     
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

     

     
     tickersymbol = Cells(2, 1).Value
     openPrice = Cells(2, 3).Value
     closePrice = Cells(2, 6).Value
     
     
     Range("J1").Value = "yearlychange"
      Range("I1").Value = "Ticker"
      Range("K1").Value = "percentageChange"
      Range("L1").Value = "totalstock"
       r = 2
       vol = 0
       
     For i = firstrow To lastRow
     
     vol = Cells(i, 7).Value + vol
     If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
     
      yearlyChange = (Cells(i, 6).Value) - (openPrice)
      
      percentageChange = (yearlyChange) / openPrice * 100
      
      openPrice = Cells(i + 1, 3).Value
      
      Range("J" & r).Value = yearlyChange
      Range("I" & r).Value = Cells(i, 1).Value
      Range("K" & r).Value = percentageChange
      Range("L" & r).Value = vol
      vol = 0
      r = r + 1
      End If
      
      Next i
      
     
     yearlyChange = Cells(i + 1, 10).Value
     
     For i = 2 To lastRow
     
     
        If Range("J" & r).Value <= 0 Then
        Range("J" & r).Interior.ColorIndex = 3
     
        ElseIf Range("J" & r).Value >= 1 Then
        Range("J" & r).Interior.ColorIndex = 4
        r = r + 1
    End If
         
     Next i
     
      
End Sub

 I used the above code to calculate the ticker, yearly change, percentage change and total volume of stock. I wrote the code for formatting the cells into green for positive and green for negative but i couldnt get it to iterate through the cells.

 
' Define variables
Dim GreatestIncrease As Double
Dim Greatestdecrease As Double
Dim GreatestTotalVolume As Double
Dim tickers As String
Dim lastRow As Long
Dim MaximumVolume As Double


        firstrow = 2
     
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("P1").Value = ticker
    Range("Q1").Value = Value
    Range("O2").Value = GreatestIncrease
    Range("O3").Value = Greatestdecrease
    Range("O4").Value = GreatestTotalVolume
    
'Loop through the worksheet and get the values needed
For i = firstrow To lastRow
    
    If Range("J" & i).Value = Maximum Then
    
        Range("P2", "Q2").Value = GreatestIncrease
        
    ElseIf Range("J" & i).Value = Minimum Then
    
        Range("P3", "Q3").Value = Greatestdecrease
        
        End If
        
    
    If Range("L" & i).Value = Maximum Then
    
    Range("P4", "Q4").Value = MaximumVolume
    
    End If
    
    Next i

    

