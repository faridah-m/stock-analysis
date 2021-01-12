# Overview of the Project 
## Project Background
* This project begun in an effort to understand how actively the stock DQ was traded. We then proceeded to do a greater analysis of the yearly return of all stock tickers across the 2017 and 2018. 

## Project Purpose
* The purpose of this project is to measure the Return across all 12 tickers in the Years 2017 and 2018. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year.
    
* In addition, the project also measures performance via a timer that displays how long it took each query, per year, to run.
   
# Results

The results across the 2017 and 2018 are as follows:
## 2017 Results:

![Image](https://github.com/faridah-m/stock-analysis/blob/main/2017_Refactored_Results.PNG)

## 2018 Results:

![Image](https://github.com/faridah-m/stock-analysis/blob/main/2018_Refactored_Results.PNG)

Full Refactored Code Below:

Sub AllStocksAnalysisRefactoredA()
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
    Dim tickerstartingPrices(12) As Single
    Dim tickerendingPrices(12) As Single
    Dim tickerVolumes(12) As Long

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
         'tickerIndex = tickers(j)
          tickerVolumes(j) = 0
       Next j
             
    '2b) Loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        For tickerIndex = 0 To 11
        If Cells(i, 1).Value = tickers(tickerIndex) Then
               tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        End If
        Next tickerIndex
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           For tickerIndex = 0 To 11
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerstartingPrices(tickerIndex) = Cells(i, 6).Value

           End If
          Next tickerIndex
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         For tickerIndex = 0 To 11
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                   
            '3d Increase the tickerIndex.
            
             tickerendingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
         End If
         Next tickerIndex
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
      For j = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + j, 1).Value = tickers(j)
        Cells(4 + j, 2).Value = tickerVolumes(j)
        Cells(4 + j, 3).Value = tickerendingPrices(j) / tickerstartingPrices(j) - 1
        
    Next j
    
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


        



Summary
