# stocks-analysis

# Purpose

The Purpose of this project was to analyze the stocks date of 2017 and 2018n and refactoring to determine whether or not the stocks are worth for investments. The Module helped me to complete the original code but to increase efficiency was the main focus point.

# The Data

The data was given to Steve for 2017 and 2018 12 different stocks analysis but same for both years. It included the the starting and closing price, date, volume and adjusted price. Zteh task for Steve was to derive results using ticker(stocks names),total daily volume and return in percentage of the stocks.


# Results

Comparison of 2017 and 2018 stocks Analysis:

First of all the original code was copied from the VBS where the refactored was required. To run both the years in one worksheet a new worksheet was created "All Stocks Analysis". In the original code some of the arrays were already created and header, range were already mentioned in it.
 
To make the code work for Steve first we needed to create ticker index to define the array for total  volumes, starting and ending prices in total and assigned to value in zero. It was required to increase the ticker volumes due to run all 12 stocks for both years.

The main and tricky part was to check the current row is the first row with selected ticker index and  same with last row and after that to increase the index by 1.

At last the output was required in three columns as tickers, volumes in total and return.

After running both years in excel Steve realized that 2017 was better in results as compare to 2018. In 2017 only stock TERP return was in negative with -7.21% as compare to 2018 where only two stocks gave the positive results ENPH 81.92% and RUN with 83.95%. All other stocks in 2018 gave the negative returns but maximum was from DQ 63% and in 2017 same stock gave the maximum return of 199.45%. That shows the portfolio of stocks has very aggressive stocks where the volatility of market is too high.



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
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i

''2b) Loop over all the rows in the spreadsheet.
         For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
     
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
      
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
                End If
   'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
                 If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
                        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
                End If
                
            '3d Increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                        tickerIndex = tickerIndex + 1
            
                End If
            
        'End If
    
        Next i
    
    
     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
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

The execution time as posted below as the run time is different for both codes could be the reason of nested loops but might be wrong.

![image](https://user-images.githubusercontent.com/79673185/111055691-b25dc900-8446-11eb-9fcc-fb4b90a5afd2.png)
![image](https://user-images.githubusercontent.com/79673185/111055697-bab60400-8446-11eb-8a8c-8ca6181d6f8c.png)


Also, the results in excel after the code ran:
2017_excel.png![image](https://user-images.githubusercontent.com/79673185/111055705-c5709900-8446-11eb-9318-e7111eadadbb.png)
2018_excel.png![image](https://user-images.githubusercontent.com/79673185/111055710-ca354d00-8446-11eb-84ef-e754ba49cce2.png)



#Summary

Refactoring helped Steve to make data look more clean and concise. With harder in programming it helped to execute the logic and write the code more effectively in terms of debugging, clean and fast improvment. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

# The Advantages of Refactoring Stock Analysis

The biggest benefit that occurred as a result of the refactoring was to decrease in macro run time. The original analysis took approximately one second to run, whereas our new analysis took for 2017 0.109 and for 2018 0.117 seconds as per attachment shown below. It could be similar every time to run but sometime can be different, it depends on data input.

  
 
   
