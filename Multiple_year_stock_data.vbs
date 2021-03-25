Sub VBS_Homework()
    
    For Each ws In Worksheets
    
    'Set variables
    Dim ticker As String
    Dim volume As Double
    Dim annual_open As Double
    Dim annual_close As Double
    
    
    ' Assign a count variable for the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
    
    ' get last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Get First Open Price
    annual_open = ws.Cells(2, 3).Value
    ws.Cells(Summary_Table_Row, 10).Value = annual_open
    
    'Create Summary Headers on each Sheet :
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Open Price"
    ws.Range("K1").Value = "Yearly Closing Price"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent Change"
    ws.Range("N1").Value = "Total Stock Volume"
    
    
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Assign Ticker values
            ticker = ws.Cells(i, 1).Value
            
            'Assign Volume values
            volume = volume + ws.Cells(i, 7).Value
    
            ' Set the First Ticker Symbol in the Summary Table when there is a change
            ws.Cells(Summary_Table_Row, 9).Value = ticker
    
            ' Set the total stock volume to the Summary Table when there is a change
            ws.Cells(Summary_Table_Row, 14).Value = volume
          
            'Get Close Price for Summary Table when there is change to the ticker symbol
            annual_close = ws.Cells(i, 6).Value
            ws.Cells(Summary_Table_Row, 11).Value = annual_close
            
            Dim Annual_change As Double
            
            ' Calculate Annual Change (Close Price - Open Price)
            Annual_change = (ws.Cells(Summary_Table_Row, 11).Value - ws.Cells(Summary_Table_Row, 10).Value)
            
                If Annual_change <> 0 And (ws.Cells(Summary_Table_Row, 10).Value <> 0) Then
                Percent_Change = FormatPercent((Annual_change / (ws.Cells(Summary_Table_Row, 10).Value)))
                Else:
                Percent_Change = FormatPercent(Annual_change)
                End If
            
            ws.Cells(Summary_Table_Row, 12).Value = Annual_change
            ws.Cells(Summary_Table_Row, 13).Value = Percent_Change
            
            'Colour Positive(Green) and Negitive(Red) Annual Change
            
            If Annual_change > 0 Then
            ws.Cells(Summary_Table_Row, 12).Interior.ColorIndex = 4 'Green for Positive Change in Stock
            ElseIf Annual_change < 0 Then
            ws.Cells(Summary_Table_Row, 12).Interior.ColorIndex = 3 'Red for Negitive Change in Stock
            End If
          
            ' Add one to the summary table row to go to the next row
             Summary_Table_Row = Summary_Table_Row + 1
          
            'Get Open Price for the next stock for the Summary Table
              annual_open = ws.Cells(i + 1, 3).Value
              ws.Cells(Summary_Table_Row, 10).Value = annual_open
             
              'reset the stock name
              ticker = ""
        
              ' Reset the Brand Total
              volume = 0
                
        ' For all same ticker symbols
        Else:
    
          ' Add to the stock total until there is a change to the ticker symbol
          volume = volume + ws.Cells(i, 7).Value
    
        End If
        
    Next i
        
    ' Summary of %Change and Volume for all stock values
        
    'Create Summary Headers on each Sheet :
    ws.Range("Q2").Value = "Greatest % Increase"
    ws.Range("Q3").Value = "Greatest % Decrease"
    ws.Range("Q4").Value = "Greatest Total Volume"
    ws.Range("R1").Value = "Ticker"
    ws.Range("S1").Value = "Value"
   
     ' Get last row for the Summary Table
     summary_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
     
     ' Calculate % Greatest Increase
     
     Dim Max_Per_Change As Double
     Dim Ticker_Max_Change As String
     
     Max_Per_Change = 0
      
     For i = 2 To summary_lastrow
     
     If ws.Cells(i, 13).Value > Max_Per_Change Then
     
     Max_Per_Change = ws.Cells(i, 13).Value
     Ticker_Max_Change = ws.Cells(i, 9).Value
     
     End If
     
     Next i
     
     'Assign Max Value with the % Format to the assigned cells
     
     Percent_Max_value = FormatPercent(Max_Per_Change)
     
     ws.Cells(2, 18).Value = Ticker_Max_Change
     ws.Cells(2, 19).Value = Percent_Max_value
     
     ' Calculate % Greatest Decrease
     
     Dim Min_Per_Change As Double
     Dim Ticker_Min_Change As String
     
     Min_Per_Change = 0
      
     For i = 2 To summary_lastrow
     
     If ws.Cells(i, 13).Value < Min_Per_Change Then
     
     Min_Per_Change = ws.Cells(i, 13).Value
     Ticker_Min_Change = ws.Cells(i, 9).Value
     
     End If
     
     Next i
     
      'Assign Min Value with the % Format to the assigned cells
     
     Percent_Min_value = FormatPercent(Min_Per_Change)
     
     ws.Cells(3, 18).Value = Ticker_Min_Change
     ws.Cells(3, 19).Value = Percent_Min_value
     
     ' Calculate Greatest Total Volume
     
     Dim Max_Tot_Volume As Double
     Dim Ticker_Max_Volume As String
     
     Max_Tot_Volume = 0
      
     For i = 2 To summary_lastrow
     
     If ws.Cells(i, 14).Value > Max_Tot_Volume Then
     
     Max_Tot_Volume = ws.Cells(i, 14).Value
     Ticker_Max_Volume = ws.Cells(i, 9).Value
     
     End If
     
     Next i
        
     ws.Cells(4, 18).Value = Ticker_Max_Volume
     ws.Cells(4, 19).Value = Max_Tot_Volume
     
    Next ws

End Sub
