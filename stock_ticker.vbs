Sub stockTicker()
  'Perform actions in each worksheet
  For Each ws in Worksheets
    
    'Add ticker, yearly change, percent change, total stock volume columns to each sheet 
    'Also adding separate ticker and value column for greatest increase/decrease/total volume
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

    'Adding var to calculate total volume of each ticker
    Dim totalVolume As Double

    totalVolume = 0

    'Adding var to determine number of rows written in totals column, set as 1 since header column was already written
    Dim numberOfTotalsWritten As Integer

    numberOfTotalsWritten = 1

    'Adding var to keep track of the first opening value of a ticker
    Dim firstRowOfTicker As Double
    'Adding var to keep track of the last opening value of a ticker
    Dim lastRowOfTicker As Double

    
    'Get last row of the worksheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1

    'Starting for loop to go through each row of worksheet
    For i = 2 To lastRow
      'Adding if statement to check if next row stock ticker is the same as current row stock ticker
      If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        'Additional if statement to check if previous row stock ticker is not the same as current stock ticker this finds the beginning of each stock ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
          'Getting opening value of the first row of the current ticker to calculate year end increase/decrease
          firstRowOfTicker = ws.Cells(i, 3)
        End If
        'Increasing totalVolume by current rows volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
      Else
        'Increasing totalVolume by current rows volume as the current row is the last row of the ticker
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        'Adding a number to numberOfTotalsWritten now to allow writing to next line
        numberOfTotalsWritten = numberOfTotalsWritten + 1
        'Getting value of the last row of the current ticker to calculate year end increase/decrease
        lastRowOfTicker =  ws.Cells(i, 6).Value
        
        'Calculating year end change by subtracking closing by opening
        ws.Cells(numberOfTotalsWritten, 10).Value = lastRowOfTicker - firstRowOfTicker
       
        'If statement to write value of the percentage increas/decrease per ticker
        If firstRowOfTicker = 0 Then
          'Warning that a the divisor is equal to 0 s
          MsgBox (ws.Cells(i, 1).Value & ": Divisor equal to 0. Setting to NaN")
          'Setting value to NaN since percentage increase from 0 is not calculable
          ws.Cells(numberOfTotalsWritten, 11).Value = "NaN"
        Else
          ws.Cells(numberOfTotalsWritten, 11).Value = Format(((lastRowOfTicker - firstRowOfTicker) / firstRowOfTicker), "Percent")
        End If
        
        'If statement to format year end change to green if greater than 0 and red if less than 0
        If ws.Cells(numberOfTotalsWritten, 10).Value > 0 Then
          ws.Cells(numberOfTotalsWritten, 10).Interior.ColorIndex = 4
        Else
          ws.Cells(numberOfTotalsWritten, 10).Interior.ColorIndex = 3
        End If
        
        'Writing ticker name to column
        ws.Cells(numberOfTotalsWritten, 9).Value = ws.Cells(i, 1).Value
        'writing totalVolume to corresponding ticker
        ws.Cells(numberOfTotalsWritten, 12).Value = totalVolume
        
        'Setting totalVolume to 0 to begin calculating total for next ticker
        totalVolume = 0
      End If

    Next i
    
    'Getting last row of results written using previous code
    lastRowOfResults = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
    'Creating vars to store highest, lowest ticker name and percentage/volume
    Dim highestPercentStockTicker As Double
    Dim highestPercentStockTickerName As String
    Dim lowestPercentStockTicker As Double
    Dim lowestPercentStockTickerName As String
    Dim highestVolumeStockTicker As Double
    Dim highestVolumeStockTickerName As String
   
    'Setting all double vars to 0 to check for higher/lower values in following code
    highestPercentStockTicker = 0
    lowestPercentStockTicker = 0
    highestVolumeStockTicker = 0

    'Starting for loop to check through each line of results starting at two due to header row
    For i = 2 To lastRowOfResults
      'If statement to check if current results percentage is higher than highestPercentageStockTicker var and also not equal to "NaN"
      If ws.Cells(i, 11).Value > highestPercentStockTicker and ws.Cells(i, 11).Value <> "NaN" Then
        'Getting the stock ticker name from column 9
        highestPercentStockTickerName = ws.Cells(i, 9).Value
        'Getting the percentage from column 11
        highestPercentStockTicker = ws.Cells(i, 11).Value
      End If
      
      'If statement to check if current results percentage is less than lowestPercentageStockTicker var and also not equal to "NaN"
      If ws.Cells(i, 11).Value < lowestPercentStockTicker and ws.Cells(i, 11).Value <> "NaN" Then
        'Getting the stock ticker name from column 9
        lowestPercentStockTickerName = ws.Cells(i, 9).Value
        'Getting the percentage from column 11
        lowestPercentStockTicker = ws.Cells(i, 11).Value
      End If

      'If statement to check if current results volume is higher than highestVolumeStockTicker var and also not equal to "NaN"
      If ws.Cells(i, 12).Value > highestVolumeStockTicker and ws.Cells(i, 11).Value <> "NaN" Then
        'Getting the stock ticker name from column 9
        highestVolumeStockTickerName = ws.Cells(i, 9).Value
        'Getting the volume from column 12
        highestVolumeStockTicker = ws.Cells(i, 12).Value
      End If

    Next i
    'Writing values to corresponding cells for highest, lowest percent/volume name and value
    ws.Cells(2, 15).Value = highestPercentStockTickerName 
    ws.Cells(2, 16).Value = Format(highestPercentStockTicker, "Percent")
    ws.Cells(3, 15).Value = lowestPercentStockTickerName
    ws.Cells(3, 16).Value = Format(lowestPercentStockTicker, "Percent")
    ws.Cells(4, 15).Value = highestVolumeStockTickerName
    ws.Cells(4, 16).Value = highestVolumeStockTicker
  
  'Do it all over again on the next worksheet
  Next ws

End Sub