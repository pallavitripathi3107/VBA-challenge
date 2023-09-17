Sub MultipleStockYear()
  
'Setting vaiables:
         Dim tickerName As String
         Dim yearlyChange As Double
         Dim percentChange As Double
         Dim closeValueForTheYear As Double
         Dim openValueForTheYear As Double
         Dim totalStockVolume As Double
         Dim lastRow As Long
         Dim greatestPercentIncrease As Double
         Dim greatestPercentDecrease As Double
         Dim greatestTotalVolume As Double
         Dim numberOfWorksheets As Integer
'wc is current worksheet
         Dim wc As Integer
'row of summary table summary table
         Dim summaryRow As Integer

'finding total worksheets in workbook
        numberOfWorksheets = Worksheets.Count
    
'Looping through all the worksheets
        For wc = 1 To numberOfWorksheets
     
'Assigning value to variable
       summaryRow = 2
       openValueForTheYear = Worksheets(wc).Cells(2, 3).Value
       totalStockVolume = 0
'Assignning Value to Greatest VAriable
      greatestPercentIncrease = 0
      greatestPercentDecrease = 0
      greatestTotalVolume = 0
      
'Adding name to the rows
        Worksheets(wc).Cells(1, 10).Value = "Ticker"
        Worksheets(wc).Cells(1, 11).Value = "Yearly Change"
        Worksheets(wc).Cells(1, 12).Value = "Percent Change"
        Worksheets(wc).Cells(1, 13).Value = "Total Stock Volume"
        
        Worksheets(wc).Cells(1, 17).Value = "ticker"
        Worksheets(wc).Cells(1, 18).Value = "value"
        Worksheets(wc).Cells(2, 16).Value = "Greatest % Increase"
        Worksheets(wc).Cells(3, 16).Value = "Greatest % Decrease"
        Worksheets(wc).Cells(4, 16).Value = "Greatest Total Volume"
        
'Changing format to column yearlyChange
       Worksheets(wc).Columns(11).NumberFormat = "0.00"
       
'Getting lastRow
      lastRow = Worksheets(wc).Cells(Rows.Count, 1).End(xlUp).Row
     
 'Loop through all tickerName
      For x = 2 To lastRow
    
     tickerName = Worksheets(wc).Cells(x, 1).Value
     totalStockVolume = totalStockVolume + Worksheets(wc).Cells(x, 7).Value
    
'Check if we are still in same ticker , if it is not...
        If Worksheets(wc).Cells(x, 1).Value <> Worksheets(wc).Cells(x + 1, 1).Value Then
           Worksheets(wc).Cells(summaryRow, 10).Value = tickerName
    
'Calculating yearly change for the year
        closeValueForTheYear = Worksheets(wc).Cells(x, 6).Value
        yearlyChange = closeValueForTheYear - openValueForTheYear
        Worksheets(wc).Cells(summaryRow, 11).Value = FormatNumber(yearlyChange, 2)
        
'Putting conditional formating on yearly change
         If yearlyChange >= 0 Then
               Worksheets(wc).Cells(summaryRow, 11).Interior.Color = vbGreen
               Worksheets(wc).Cells(summaryRow, 12).Interior.Color = vbGreen
          Else
               Worksheets(wc).Cells(summaryRow, 11).Interior.Color = vbRed
               Worksheets(wc).Cells(summaryRow, 12).Interior.Color = vbRed
         End If
         
'Calculating Percentage Change
         percentChange = (yearlyChange / openValueForTheYear)
        
'Assigning Percent change to summary table
         Worksheets(wc).Cells(summaryRow, 12).Value = Format(percentChange * 100, "#.00") + "%"
         
 'Calculating Greatest Percentage Change
        If percentChange > greatestPercentIncrease Then
            greatestPercentIncrease = percentChange
            Worksheets(wc).Cells(2, 18).Value = Format(greatestPercentIncrease * 100, "#.00") + "%"
            Worksheets(wc).Cells(2, 17).Value = tickerName
        End If
        
        If percentChange < greatestPercentDecrease Then
            greatestPercentDecrease = percentChange
            Worksheets(wc).Cells(3, 18).Value = Format(greatestPercentDecrease * 100, "#.00") + "%"
            Worksheets(wc).Cells(3, 17).Value = tickerName
        End If
    
        If totalStockVolume > greatestTotalVolume Then
           greatestTotalVolume = totalStockVolume
           Worksheets(wc).Cells(4, 18).Value = greatestTotalVolume
           Worksheets(wc).Cells(4, 17).Value = tickerName
        End If
    
 'Assigning totalStockVolume value
        Worksheets(wc).Cells(summaryRow, 13).Value = totalStockVolume
        
'This open value is for next ticker ex.(AAF)
        openValueForTheYear = Worksheets(wc).Cells(x + 1, 3)
        
 'Row for summary table adding one row
        summaryRow = summaryRow + 1
        
'Resetting total stock volume for next ticker
        totalStockVolume = 0
        
       End If
    
      Next x
    
  Next wc

End Sub
