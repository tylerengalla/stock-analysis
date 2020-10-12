# stock-analysis

# Overview of Project

Steve needs help to analyze stock performances similar to the ones his parents are interested in investing in - particularly around Green Energy Stocks.  

We're going to use VBA to help us understand how Daqo (the company's stock Steve's parents are interested in) compares in terms of total daily volume (or shares traded throughout the day - measuring how active the stock is) and the yearly return for each stock. 

But we also want to be able to use these macros for future analysis if needed - so we'll create buttons to help them run this analysis and understand if this or any other stock would be a good investment.  


# Results

From running our macros to return the Total Daily Volume and the Return for our stocks we can see how DQ is doing compared to the other stocks. 

In 2017 - Daqo (DQ) was not heavily traded compared to the other stocks but yielded the highest return at 199.4%. 

![](/Stock_Performance_2017.png)

Below you can see how fast we were able to run this subroutine:

![](/VBA_Challenge_2017.png)

In 2018 - Daqo (DQ) tripled it's Total Daily Volume and yielded a negative return of -62.6%. Noticeably though, all but two of the stocks yielded negative returns meaning the industry as a whole is on a downward trend as compared to the previous year where only 1 stock yielded negative returns. This could be a good time to buy. 

![](/Stock_Performance_2018.png)

Below you can see how fast we were able to run the subroutine on the second year. 

![](/VBA_Challenge_2018.png)


# Summary
* Advantages of refactoring code is that you're not starting from scratch and already have a solution to the problem you're looking to solve. You just have to optimize it now. 
* Disadvantages is that you could potentially be following somebody else's code logic and have to "get up to speed" and be hopeful that they provided good enought comments to follow their thread. 

Luckily we were part of the original VBA so we were familiar, but I can see that not always being the case. 

# Code Used

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
    
      Dim tickerIndex As Single
    
      tickerIndex = 0


      '1b) Create three output arrays
      Dim tickerVolume(12) As Long
      Dim tickerStartingPrice(12) As Single
      Dim tickerEndingPrice(12) As Single
    
    
      ''2a) Create a for loop to initialize the tickerVolumes to zero.
      ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
      For i = 0 To 11
        
          tickerVolume(i) = 0
        
      Next i
    

      ''2b) Loop over all the rows in the spreadsheet.
    
          For i = 2 To RowCount
    
              '3a) Increase volume for current ticker
        
              tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
            
        
              '3b) Check if the current row is the first row with the selected tickerIndex.
              'If  Then
         
              If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         
                  tickerStartingPrice(tickerIndex) = Cells(i, 6).Value

              End If
        
            
          'End If
        
          '3c) check if the current row is the last row with the selected ticker
          'If  Then
        
              If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
                  tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            

              '3d Increase the tickerIndex.
              tickerIndex = tickerIndex + 1

              End If
            
          'End If
    
          Next i
    
      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11
        
          Worksheets("All Stocks Analysis").Activate

          Cells(4 + i, 1).Value = tickers(i)
          Cells(4 + i, 2).Value = tickerVolume(i)
          Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
        
        
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
    
    Sub ClearWorksheet()

      Cells.Clear
    
    End Sub


