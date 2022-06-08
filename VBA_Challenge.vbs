Sub AllStocksAnalysisRefactored()
    Dim starttime As Single
    Dim endtime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

        starttime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
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
    Sheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerindex = 0
    

    '1b) Create three output arrays
    Dim tickervolumes(12) As Long
    Dim tickerstartingprices(12) As Single
    Dim tickerendingprices(12) As Single

    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        ticker = tickers(i)
        tickervolumes(UBound(tickervolumes)) = 0
    
        
        
    ''2b) Loop over all the rows in the spreadsheet.
    Sheets(yearvalue).Activate
    For c = 2 To RowCount
    
        '3a) Increase volume for current ticker
         If Cells(c, 1).Value = ticker Then
        
            tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(c, 8).Value
            
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(c - 1, 1).Value <> ticker And Cells(c, 1).Value = ticker Then
            'set starting price
            tickerstartingprices(tickerindex) = Cells(c, 6).Value
        End If
        
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(c + 1, 1).Value <> ticker And Cells(c, 1).Value = ticker Then
            'set ending price
            tickerendingprices(tickerindex) = Cells(c, 6).Value
        End If
            

            '3d Increase the tickerIndex.
            If Cells(c + 1, 1).Value <> ticker And Cells(c, 1) = ticker Then
            tickerindex = tickerindex + 1
            End If
            
        'End If
    Next c


    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
          
  
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Cells(i + 4, 1).Value = tickers(i)
    Cells(i + 4, 2).Value = tickervolumes(i)
    Cells(i + 4, 3).Value = tickerendingprices(i) / tickerstartingprices(i) - 1
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
Next i
    dataRowStart = 4
    dataRowEnd = 15

    For x = dataRowStart To dataRowEnd
        
        If Cells(x, 3) > 0 Then
            
            Cells(x, 3).Interior.color = vbGreen
            
        Else
        
            Cells(x, 3).Interior.color = vbRed
            
        End If
        
    Next x
 
    endtime = Timer
    MsgBox "This code ran in " & (endtime - starttime) & " seconds for the year " & (yearvalue)
End Sub
Sub Anyyearnalysis()
    'define variable type
    Dim starttime As Single
    Dim endtime As Single
    
    'input year you want analysis run
    yearvalue = InputBox("What year would you like to run the analysis on?")
        
        starttime = Timer
        '1)Format the output sheet on the "All Stocks Analysis" worksheet.
    
    Worksheets("all stocks analysis").Activate
        Range("A1").Value = "All Stocks (" + yearvalue + ")"
        
        'create a header row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
    '2)Initialize an array of all tickers.
    
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
        
    '3)Prepare for the analysis of tickers.
        '3a)Initialize variables for the starting price and ending price.
        
        Dim startingPrice As Single
        Dim endingPrice As Single
        
        '3b)Activate the data worksheet.
        
        Sheets(yearvalue).Activate
        
        '3c)Find the number of rows to loop over.
        
        RowCount = Cells(Rows.Count, "a").End(xlUp).Row
        
    '4)Loop through the tickers.
    
    For i = 0 To 11
        
        ticker = tickers(i)
        totalvolume = 0
        
    '5)Loop through rows in the data.
    
    Sheets(yearvalue).Activate
    For j = 2 To RowCount

    
        '5a)Find the total volume for the current ticker.
        
        If Cells(j, 1).Value = ticker Then
        
            totalvolume = totalvolume + Cells(j, 8).Value
            
        End If
        
        '5b)Find the starting price for the current ticker.
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            'set starting price
            startingPrice = Cells(j, 6).Value
        End If
        
        '5c)Find the ending price for the current ticker.
        
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            'set ending price
            endingPrice = Cells(j, 6).Value
        End If
        
    Next j
    
    '6)Output the data for the current ticker.
        
        Worksheets("all stocks analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalvolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        Columns("B").AutoFit
        Range("b4:b15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.00%"
   Next i
      dataRowStart = 4
    dataRowEnd = 15
       For x = dataRowStart To dataRowEnd
        
        If Cells(x, 3) > 0 Then
            
            Cells(x, 3).Interior.color = vbGreen
            
        Else
        
            Cells(x, 3).Interior.color = vbRed
            
        End If
        
    Next x
   endtime = Timer
   MsgBox "this code ran in " & (endtime - starttime) & "second for the year" & (yearvalue)
   
End Sub