# Utilizing VBA Macros to gather year long returns on stock portfolio. 

## Overview
We want to help Steve by improving our stock anaylsis code to be more scalable and to potentially include any number of stock tickers rather than the original 12 supplied while simultaneously increasing the efficciency with with the code runs. 

## Results
With fewer If statements and no need for any nested For statments we were able to produce the same outcome in a significantly shorter amount of time.<br><br>
<img src="https://user-images.githubusercontent.com/15967377/163686798-36d4d023-570c-47d6-9ffd-2c00b4a8b2da.PNG" width="425" height="275" title="Original 2017">
<img src="https://user-images.githubusercontent.com/15967377/163686804-50611b28-a1f3-447b-8a2b-068e588f6a36.PNG"   width="425" height="275" title="Refactored 2017">
<img src="https://user-images.githubusercontent.com/15967377/163687902-8578af3b-f809-4e12-9eab-a16f0d1573e0.PNG"  width="425" height="275" title="Original 2018">
<img src="https://user-images.githubusercontent.com/15967377/163687905-8c238c79-5bd0-430d-bdd8-ba77ee7f6bc8.PNG" width="425" height="275" title="Refactored 2018"><br><p>
            
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

We also used arrays to store our ticker volume, starting and ending price which allowed us to quickly loop over all when it came time to display them. 
```vba
For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

    Next i
```

## Summary
*1. What are the advantages or disadvantages of refactoring code?*<br>
Refactoring code can make it more effcient as well improving readability in turn making it easier to maintain or modify. 
It does take time to refactor code, whether it is worth the time seems as though it depends on the case in which it is being applied. Keep in mind that bugs are more likely to be caught when refactoring and that alone could prove to be enough of a reason to make refactoring good standard practice.

*2. How do these pros and cons apply to refactoring the original VBA script?*<br>
It definitely took me some time and many steps through while debugging and tracking the outcomes of each loop to ensure the correct outcome was reached but with a reduction of run
