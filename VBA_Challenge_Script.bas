Attribute VB_Name = "Module1"
Sub YearlyStockData():

' Defining variables for the sheets.
'---------------------------------------------------------------------------------
Dim ticker As String
Dim number_tickers As Integer
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double

Dim lastRowState As Long

'variables for bonus portion
'------------------------------------------------------------------------------------
'Greatest % Increase
Dim greatest_increase As Double
Dim greatest_increase_ticker As String

'Greatest % Decrease
Dim greatest_decrease As Double
Dim greatest_decrease_ticker As String

'Greatest Stock Volume
Dim greatest_stockvolume As Double
Dim greatest_stockvolume_ticker As String


' Initiate loop over each sheet in the workbook.
'----------------------------------------------------------------------------------
For Each ws In Worksheets

ws.Activate

'Go to the last row.

lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Headers for specified columns for each sheet.

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
    
'Variables for each sheet.

number_tickers = 0
ticker = ""
yearly_change = 0
opening_price = 0
percent_change = 0
total_stock_volume = 0
    
'Initiate loop through the list of tickers.


For i = 2 To lastRowState

    ticker = Cells(i, 1).Value
        
If opening_price = 0 Then
   opening_price = Cells(i, 3).Value
End If
        
total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
'Run this if a different ticker is displayed.

    If Cells(i + 1, 1).Value <> ticker Then
        number_tickers = number_tickers + 1
        Cells(number_tickers + 1, 9) = ticker
            
        closing_price = Cells(i, 6)
            
        yearly_change = closing_price - opening_price
        Cells(number_tickers + 1, 10).Value = yearly_change
'------------------------------------------------------------------------------
'Color Formatting
    If yearly_change > 0 Then
        Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
    ElseIf yearly_change < 0 Then
           Cells(number_tickers + 1, 10).Interior.ColorIndex = 3

    Else
              
    End If
            
            
    If opening_price = 0 Then
       percent_change = 0
    Else
        percent_change = (yearly_change / opening_price)
    End If

'Formatting for percent
Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")


'Reset opening price back to 0 when a different ticker is displayed.

opening_price = 0
Cells(number_tickers + 1, 12).Value = total_stock_volume
            
'Reset total stock volume back to 0 when a different ticker is displayed.
            total_stock_volume = 0
    End If
        
    Next i
    
'Assigning names for greatest increase, greatest decrease, and greatest volume.
'---------------------------------------------------------------------------------------
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Find last row
lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row

'Values of the bonus portion variables
greatest_increase = Cells(2, 11).Value
greatest_increase_ticker = Cells(2, 9).Value
greatest_decrease = Cells(2, 11).Value
greatest_decrease_ticker = Cells(2, 9).Value
greatest_stockvolume = Cells(2, 12).Value
greatest_stockvolume_ticker = Cells(2, 9).Value

'-------------------------------------------------------
'Initiate loop for list of tickers/ bonus portion.

For i = 2 To lastRowState

If Cells(i, 11).Value > greatest_increase Then
    greatest_increase = Cells(i, 11).Value
    greatest_increase_ticker = Cells(i, 9).Value
End If

If Cells(i, 11).Value < greatest_decrease Then
    greatest_decrease = Cells(i, 11).Value
    greatest_decrease_ticker = Cells(i, 9).Value
End If

If Cells(i, 12).Value > greatest_stockvolume Then
    greatest_stockvolume = Cells(i, 12).Value
    greatest_stockvolume_ticker = Cells(i, 9).Value
End If


Next i

'Add the loop to each sheet
Range("P2").Value = Format(greatest_increase_ticker, "Percent")
Range("Q2").Value = Format(greatest_increase, "Percent")
Range("P3").Value = Format(greatest_decrease_ticker, "Percent")
Range("Q3").Value = Format(greatest_decrease, "Percent")
Range("P4").Value = greatest_stockvolume_ticker
Range("Q4").Value = greatest_stockvolume
    
    
Next ws


End Sub
