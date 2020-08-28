Sub SM_Analyzer()

    'Set name for cells:
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Set Variable for Ticker symbol,yearly change, percent change, and total stock vol
    Dim ticker_sym As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As Double
    
    'Set variable for ticker start and ticker end (close)
    Dim yearclose_first As Double, yourclose_last As Double
    
    'Set variable for year change value
    Dim yearchange_value As Double
    
    'Keep track of ticker in table
    Dim ticker_summary_row As Integer
    ticker_summary_row = 2
    
    'Look through all ticker
    For i = 2 To 70926
    
    'Define initial value of day close for ticker
   If yearclose_first = 0 Then
    yearclose_first = Cells(i, 6).Value
    
    End If
    
    'Define initial value for year change
      If yearchange_value = 0 Then
    yearchange_value = Cells(i, 10).Value
    
    End If
    
    
    'Check if ticker is still the same
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i - 1, 1) = Cells(i, 1) Then
    
    'Define last close per ticker
    yearclose_last = Cells(i, 6).Value
    
    'Yearly Change
    yearly_change = yearclose_last - yearclose_first
    
    'Percent Change
    percent_change = yearchange_value / yearclose_first
    
    'Set ticker symbol
    ticker_sym = Cells(i, 1).Value
    
    'Add to total stock
    total_stock = total_stock + Cells(i, 7).Value
    
    
    ' Add ticker to table
      Range("I" & ticker_summary_row).Value = ticker_sym
      
    'Add yearly change to table
        Range("J" & ticker_summary_row).Value = yearly_change
        'format conditional color (red if negative, else green)
        If Range("J" & ticker_summary_row).Value < 0 Then
         Range("J" & ticker_summary_row).Interior.ColorIndex = 3
         Else
          Range("J" & ticker_summary_row).Interior.ColorIndex = 4
        End If
        
    'Add percent change to table
        Range("K" & ticker_summary_row).Value = percent_change
        'format to percent
        Range("K" & ticker_summary_row).NumberFormat = "0.00%"
 
    'Add total stock to table
      Range("L" & ticker_summary_row).Value = total_stock

    ' Add to ticker table
      ticker_summary_row = ticker_summary_row + 1
      
      ' Reset the total stock, the yearly change start, and the percent change
      total_stock = 0
      yearclose_first = 0

  
    Else

      ' Add to the stock total
      total_stock = total_stock + Cells(i, 7).Value
      
    yearly_change = yearclose_last - yearclose_first

    End If
    
    Next i
    
    
    
    
    
End Sub
