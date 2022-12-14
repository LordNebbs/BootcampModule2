# Stock Analysis with VBA

## Overview of Project
We have been tasked with creating a stock price analysis for our friend Steve. He has been given access to the stock data for several companies and how they performed in 2017 and 2018. Steve needs us to go through the data and analyze and create a digestible version of the data for 12 green energy companies to make financial decisions moving forward.
 
## Results
- By filtering the data from dozens of stocks down to the stock tickers Steve needs, we are able to determine the Total Daily Volume and the annual return percentage.

### Project plan
While we could filter the data manually, we determined to have Excel tabulate the data automatically with the click of a button.
- First we established an arrary containing the Stock Tickers
```
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"


    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
-Next we set up For Loops to process the stock data

    For i = 2 To RowCount
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        If the next rows ticker doesnt match, increase the tickerIndex.
        
        If Cells(i, 1).Value <> tickers(tickerIndex) And Cells(i + 1, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        If Cells(i, 1).Value <> tickers(tickerIndex) And Cells(i + 1, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
        
    Next i
    

-Once the macros have been created to fetch the data and display it in a more digestible form, we need to format the data in a way that allows the user to interpret it clearly as well as making buttons to reset and run the calculation.

    For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

    Next i


    Sub DQAnalysis_Click()
    DQAnalysis
    DQFormatting
    End Sub

    Sub AllAnalysis_Click()
    AllStocksAnalysis
    ALLFormatting
    End Sub
    
    Sub ClearButton_Click()
    ClearWorksheet
    End Sub
```
Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 ```   

Using the above Macros we were able to produce the following tables. 

![2017](https://github.com/LordNebbs/BootcampModule2/blob/main/VBA_Challenge_2017.png) ![2018](https://github.com/LordNebbs/BootcampModule2/blob/main/VBA_Challenge_2018.PNG)

- From the data we were able to determine that Green Energy companies did very well in 2017 but lost gains the following year. Our advice to Steve would be that RUN and ENPH seem to be the best investment choice as they continue to have near 100% gains when the entire sector suffered losses. If Steve wanted more risk the recommendation would be investing in SEDG and VSLR because while they did show losses, they only suffered single digit percentage losses off of large gains the previous year compared to the rest of the sector. 

- Using this data, Steve has an excellent resource to make informed investment decisions.

## Summary
The refactored version of the code was simpler and perhaps cleaner. Having a completed outline makes completing tasks easier but you are completely dependent on the direction that outline is giving you, compared to the original where you could take the code in any direction you wanted at the cost of complexity (perhaps needlessly)
