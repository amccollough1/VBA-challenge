Attribute VB_Name = "Module1"
Sub Year_Stock()
  For Each ws In Worksheets
    Dim WorksheetName As String
    Dim ticker As Integer
    Dim beginning As Double
    Dim closing As Double
    Dim Volume As Double
    Dim GreatestInc As Double
    Dim GreatestDec As Double
    Dim GreatestVolume As Double
    
    WorksheetName = ws.Name
    'Label Titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest%Increase"
    ws.Cells(3, 15).Value = "Greatest%Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    
    'Finds last row in  <ticker> column
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'set ticker value
    ticker = 2
    'start row
     r = 2
     
    'Loop through all rows
    For i = 2 To LastRow
    

   'condition to check if ticker symbol changed
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'writes ticker symbol  in column I
        ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
   
        
        'calculate yearly change in price  from the opening price at the beginning of a given year to the closing price at the end of that year
         beginning = ws.Cells(r, 3).Value
         
         LastValue = ws.Cells(i + 1, 1).Row - 1
         closing = ws.Cells(LastValue, 6).Value
         Change = closing - beginning
         
        'Writes difference of change in column J
         ws.Cells(ticker, 10).Value = Change
         
         'Conditional Format to highlight positive change in green and negative in red
         If ws.Cells(ticker, 10).Value < 0 Then
            ws.Cells(ticker, 10).Interior.ColorIndex = 3
         Else
            ws.Cells(ticker, 10).Interior.ColorIndex = 4
         End If
         
         
         'Percentage Change  from the opening price at the beginning of a given year to the closing price at the end of that year
         PercentChange = (Change / beginning)
         
         'Writes value of PercentChange in column K
         ws.Cells(ticker, 11).Value = PercentChange
         
         'Calculate total stock volume of the stock
        
         ws.Cells(ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(r, 7), ws.Cells(i, 7)))
         
      
         
         
         'Add ticker +1
         ticker = ticker + 1
         'move to next row
         r = i + 1
    
    
        
   End If
        
    
    Next i
    
    
    'Find Greatest values
    
    LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

   'Loop through all of row i, where the ticker values are listed
   For i = 2 To LastRowI
        
     'Find Greatest Increase %
    If ws.Cells(i, 11).Value > GreatestInc Then
        GreatestInc = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    Else
        GreatestInc = GreatestInc
    End If
    
    
    'Find Greatest Decrease %
    If ws.Cells(i, 11).Value < GreatestDec Then
        GreatestDec = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    Else
        GreatestDec = GreatestDec
    End If
        
        
        
     'Find Greatest Volume
    If ws.Cells(i, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    Else
        GreatestVolume = GreatestVolume
    
    End If
     
     
     Next i
     


    
   
  Next ws
  
    
    
End Sub
