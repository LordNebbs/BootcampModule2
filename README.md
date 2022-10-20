# Stock Analysis with VBA

## Overview of Project
We have been tasked with creating a stock price analysis for our friend Steve. He has been given access to the stock data for several companies and how they performed in 2017 and 2018. Steve needs us to go through the data and analyze and create a digestible version of the data for 12 green energy companies to make financial decisions moving forward.
 
## Results
- By filtering the data from dozens of stocks down to the stock tickers Steve needs, we are able to determine the Total Daily Volume and the annual return percentage.
- From the data we were able to determine that Green Energy companies did very well in 2017 but lost gains the following year. Our advice to Steve would be that RUN and ENPH seem to be the best investment choice as they continue to have near 100% gains when the entire sector suffered losses. If Steve wanted more risk the recommendation would be investing in SEDG and VSLR because while they did show losses, they only suffered single digit percentage losses off of large gains the previous year compared to the rest of the sector. 
![2017](https://github.com/LordNebbs/BootcampModule2/blob/main/VBA_Challenge_2017.png) ![2018](https://github.com/LordNebbs/BootcampModule2/blob/main/VBA_Challenge_2018.PNG)

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


Using this data, Steve has an excellent resource to make informed investment decisions.

### Summary

