# Stocks-Analysis
Analyzing stock options 

I.Overview of the Project
   
    The project is regarding the ability to use VBA in analyzing multiple stocks which are very crucial within the financial industries. The idea here is to write codes that will automate the analyses of stocks for each years. Using codes to automate tasks helps decrease one's chances of errors and reduces the time needed to run analyses. So for the VBA challenge project, VBA coding was used to collect stock information for 2017 and 2018 in order to determine whether stocks were worth investing for the following years. 

II. Results

    The data includes two separate charts, one from the year 2017 and one from the year 2018. Each chart consists of 12 different stocks which contain ticker of the stock, the total daily volume, and the return. In order to get the following information, the VBA coding for determining ticker value, the year of the stock, determining the highest and lowest price of stocks, along with the volume of stocks were typed and run.

III. Summary

    Advantages and Disadvantages of Refactoring Code in General:
    The advantages of refactoring code are including: design and software improvement, efficient programmng, and having a more organized set of codes which helps with debugging. The disadvantages may include not having proper test runs for the set of codes.

    Advantages and Disadvantages of the Original and Refactored VBA Script
    The advantages of the original VBA Script and Refactored VBA Script include more organized coding, giving a simpler method of understanding how to include loops and determine the following prices. Disadvantages include the amount of time it took to recieve the following analyses for each year. Here are the following pics provided for both 2017 and 2018 along with the coding.





    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerindex = 0

    '1b) Create three output arrays
    Dim tickerVolume As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11:
        ticker = tickers(i)
        tickerVolume = 0
    
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    Worksheets(yearValue).Activate
    
        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = ticker Then
            tickerVolume = tickerVolume + Cells(j, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            tickerStartingPrices = Cells(j, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            tickerEndingPrices = Cells(j, 6).Value
        
        End If
        
    
    Next j


    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = tickerVolume
        Cells(4 + 1, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
        
    Next i
    
    'Formatting
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub





