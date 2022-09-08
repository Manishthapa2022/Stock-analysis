# **Stock-analysis**
## Overview of the Project
The main purpose o the project is to provide Steve the Total daily Volume and the Yearly return for a portfolio of stocks so that he can advise his parents to invest accordingly for better returns. The analysis was carried out both using a original and a refactored code. The refactored code was used specifically so that Steve could add more data sets to it which he could not to the original code. 
## Results:
The following illustrates for the portfolio of stocks using Original and Refactored code for the year 2018.  

##### **Original code**
 ~~~VBAscript
 Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single
    Dim yearValue As String
    Dim tickerIndex As Integer
    Dim RowCount As Long
    Dim currentRow As Long
    Dim currentTicker As String
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer

    'Initialize array of all tickers
    Dim tickers(11) As String
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index, initialize tickerIndex
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
        tickerStartingPrices(tickerIndex) = 0
        tickerEndingPrices(tickerIndex) = 0
    Next tickerIndex

    ''2b) Loop over all the rows in the spreadsheet.
    Sheets(yearValue).Activate
    tickerIndex = 0
    For currentRow = 2 To RowCount
        currentTicker = tickers(tickerIndex)

        '3a) Increase volume for current ticker
        If Cells(currentRow, 1).Value = currentTicker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(currentRow, 8).Value
        Else
            MsgBox "Error: ticker mismatch" + CStr(currentRow) + currentTicker
            Exit Sub
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(currentRow - 1, 1).Value <> Cells(currentRow, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(currentRow, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(currentRow + 1, 1).Value <> Cells(currentRow, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(currentRow, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next currentRow
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For tickerIndex = 0 To 11
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
        Cells(4 + tickerIndex, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + tickerIndex, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    Next tickerIndex
    
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Format the table header, number formats, and column B
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    Const DATA_START As Integer = 4
    Const DATA_END As Integer = 15

    Dim returnPercent As Integer

    'Color the background of all negative returns red and all positive returns green
    For returnPercent = DATA_START To DATA_END
        If Cells(returnPercent, 3) > 0 Then
            Cells(returnPercent, 3).Interior.Color = vbGreen
        Else
            Cells(returnPercent, 3).Interior.Color = vbRed
        End If
    Next returnPercent
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

 ~~~
![2018 Green Stock analysis with Run time](https://github.com/Manishthapa2022/Stock-analysis/blob/main/Resource/Green_Stocks_2018.png)
##### **Refactored code**
![2018 VBA Challenge with Run time](https://github.com/Manishthapa2022/Stock-analysis/blob/main/Resource/VBA_Challenge_2018.PNG)
---
Based on the above, the following observations were made
- ENPH at 81.9% and RUN a 84% were the only two stocks which had the positive returns whereas JKS at 60.5% and DQ at 62.6% had the lowest return for 2018.
- The total run time for refactored code for 2018 was 0.0390625 secs and was 0.3242188 secs for the original code.
--- 
## Summary:
### Advantages and Disadvantages of using the Refactored code
The advantages of using the refactored code are that it can be used on a much larger data set and the time taken for the execution of the code is relatively smaller, whereas on the other hand lot of information needs to be added including additional arrays and variables because of the larger data sets that eventually can corrupt the previous code.   
### Advantages and Disadvantages of the Original and Refactored code
The Original code was much simplier and we could easil obtain the Summations of Daily volumes and Retunrs using the current Data Set provided by Steve, whereas the run time was much more at 0.3242188 secs for 2018. Also, we could not carry out much detailed analysis as only Tickers were used as ann array whereas with refactored code, Ticker volumes, Tickerstarting prices and TickerEnding prices were also used. 
The advantages of using the refactored code are that the run time was much smaller at 0.0390625 secs (2018 data), and also Steve can add thousands of Data sets which would not be possible with the original code. Although it can be very time consuming espcially when the code is long and complicated. Good coding knowledge and understanding of the data is crucial as modifying the original code can corrupt it. 
