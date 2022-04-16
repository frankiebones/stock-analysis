# Utilizing VBA Macros to gather year long returns on stock portfolio. 

## Overview
We want to help Steve by refactoring our stock anaylsis code to be more scalable and to potentially include any number of stock tickers rather than the original 12 supplied. 

## Results
With less If statements and no need for any nested For statments we were able to produce the same outcome in a significantly shorter amount of time.

We find the starting price and the ending price of a particular ticker with a series of If statements by comparing the iterator of the loop to the current index of the tickerIndex variable which we then increase by 1 once we confirm that we have arrived at the last row containing the current ticker.
```vba
If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
  tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
End If
        
If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
  tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

  tickerIndex = tickerIndex + 1

End If
```

We also used arrays to store our ticker volume, starting and ending price which allowed us to quickly loop over them when it came time to display them. 
```vba
For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

    Next i
```

## Summary
1. What are the advantages or disadvantages of refactoring code?

2. How do these pros and cons apply to refactoring the original VBA script?
