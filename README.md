# Stock-Analysis

## Overview of Project
A close friend, Steve, has recently graduated with a fininace degree and he is looking for help in analyzing green energy stocks for his parents who are passinoate about the future of renewable energy. Using Visual Basic for Applications through Excel, an analysis is performed on "Daqo" and then continued to investigate 12 stocks in total. 

### Purpose
The purpose of this challenge is to refactor VBA code previously used to analyze the Stock Market Dataset from 2017 and 2018, in order to make a faster analysis. Refactoring code involves restructuring existing code without changing it's output. By taking fewer step, using less memory, and improving the logic of our code, we can not only make it easier to read but make it more efficient. The main objective is to provide Steve with the most efficient code to execute future analyses in larger data sets. 


## Results
### Original Analysis 
```
Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    '1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    
    'Make the title in cell A1 "All Stocks (2018)"
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Add three columns with headers: ticker, total daily volume, return
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '2) Initialize an array of all the tickers
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
    
    '3a) Initialize variables for the starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    '3b) Activate the data worksheet
    Worksheets(yearValue).Activate
    
    '3c) Find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through the tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
    '5) Loop through rows in the data
    Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
    '5a) Find total volume for the current ticker
        If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
    
    '5b) Find starting price for the current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            
        End If
        
    '5c) Find ending price for the current ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            
        End If
        
        Next j
        
    '6) Output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i

End Sub
```

### Refactored Analysis
```
 tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```


    '1a) Create a ticker Index
    Dim tickerIndex As Single
        tickerIndex = 0
    

    '1b) Create three output arrays
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    
    Next i
    
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
    '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
    End If
        
    '3c) check if the current row is the last row with the selected ticker
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
          tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        '3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
            
    End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
       
    Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(4 + i, 1).Value = tickers(tickerIndex)
    Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
    Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
        
    Next i
### Elapsed Time: Original Analysis
![image](resources/Theater_Outcomes_vs_Launch.png) 
### Elapsed Time: Refactored Analysis

## Summary
### Advatages and Disadvantages of Refactoring Code
The main advantages of refactoing code is creating a script that is easier to read, but more importantly having a more efficient code. Efficiency in a script will make it possibile to analyze and collect data faster. A disadvantage is ruining an already successful code. Taking what is already capable of performing an analysis and making it better is the main objective in this challenge, however it is essential that it is refactored correctly.  

### Original Script Versus Refactored Script
In this particular project it was beneficial to refactor the original code. It will allow further analyses be performed on datasets that are much larger holding potentially thousands of stocks as compared to our dataset with only twelve. The diffiuclt part in refactoring this code is comprehending the differences in syntax between the origninal and refactored scripts. Without the proper usage of VBA language, the refactoring can take much longer than desired. 
