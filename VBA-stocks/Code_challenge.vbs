Sub Final_test_2()

' Defining The Different Variables
Dim Opening_price As Double
Dim Closing_price As Double
Dim Difference As Double
Dim Percent_change As Double
Dim cumulative_vol As Double
Dim current_vol As Double
Dim greatest_percent As Double
Dim lowest_percent As Double
Dim stock_volume As Double
Dim ticker_val_high As String
Dim ticker_val_low As String
Dim ticker_val_stock As String


For Each ws In Worksheets
' Inserting the headers of the summary table
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"
ws.Range("J1:M1").Font.Bold = True


Let i = 2
For r = 1 To 797711

' Obtaining the Ticker and the opening price every time there is a new ticker stock
If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
ws.Cells(i, 10).Value = ws.Cells(r + 1, 1).Value
End If

' Tallying up the volume
current_vol = ws.Cells(r + 1, 7).Value
cumulative_vol = cumulative_vol + current_vol

'Obtaining the opening Price for Stocks
If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value And ws.Cells(r + 1, 3).Value <> 0 Then
Opening_price = ws.Cells(r + 1, 3)
' Adjusting the opening price if it starts of as a 0 and eventually becomes a different opening price on a different day
ElseIf ws.Cells(r + 1, 1).Value = ws.Cells(r, 1).Value And ws.Cells(r, 3).Value = 0 And ws.Cells(r + 1, 3).Value <> 0 Then
Opening_price = ws.Cells(r + 1, 3).Value
' Opening price if it is  0 the whole way through
ElseIf ws.Cells(r + 1, 1).Value = ws.Cells(r, 1).Value And ws.Cells(r, 3).Value = 0 And ws.Cells(r + 1, 3).Value = 0 Then
Opening_price = 0
End If

'Obtaining the Closing price for Stocks
If ws.Cells(r + 2, 1).Value <> ws.Cells(r + 1, 1).Value Then
Closing_price = ws.Cells(r + 1, 6)
'Avoiding the Dividing by 0 problem
If Opening_price = 0 And Closing_price = 0 Then
Percent_change = 0
Difference = 0
Else
Difference = Closing_price - Opening_price
Percent_change = Difference / Opening_price
End If
ws.Cells(i, 12).Value = Percent_change
' Conditional Formatting for the Yearly column
ws.Cells(i, 11).Value = Difference
If Difference < 0 Then
ws.Cells(i, 11).Interior.ColorIndex = 3
Else
ws.Cells(i, 11).Interior.ColorIndex = 4
End If

ws.Cells(i, 13).Value = cumulative_vol

cumulative_vol = 0

i = i + 1

End If

Next r
 
'Challenge Activity 

 greatest_percent = 0
lowest_percent = 0
stock_volume = 0
ticker_val_high = "A"
ticker_val_low = "A"
ticker_val_stock = "A"


'Obtaining Greatest percent and the ticker value associated 
For j = 2 To 3000
If ws.Cells(j + 1, 12).Value > ws.Cells(j, 12).Value And ws.Cells(j + 1, 12).Value > greatest_percent Then
greatest_percent = ws.Cells(j + 1, 12).Value
ticker_val_high = ws.Cells(j + 1, 10).Value
ElseIf ws.Cells(j, 12).Value > ws.Cells(j + 1, 12).Value And ws.Cells(j, 12).Value > greatest_percent Then
greatest_percent = ws.Cells(j, 12).Value
ticker_val_high = ws.Cells(j, 10).Value
Else
greatest_percent = greatest_percent
ticker_val_high = ticker_val_high
End If


'Obtaining lowest percent and the ticker value associated 
If ws.Cells(j + 1, 12).Value < ws.Cells(j, 12).Value And ws.Cells(j + 1, 12).Value < lowest_percent Then
lowest_percent = ws.Cells(j + 1, 12).Value
ticker_val_low = ws.Cells(j + 1, 10).Value
ElseIf ws.Cells(j, 2).Value < ws.Cells(j + 1, 12).Value And ws.Cells(j, 12).Value < lowest_percent Then
lowest_percent = ws.Cells(j, 12).Value
ticker_val_low = ws.Cells(j, 10).Value
Else
lowest_percent = lowest_percent
ticker_val_low = ticker_val_low
End If

'Obtaining Highest stick volume and the ticker value associated 
If ws.Cells(j + 1, 13).Value > ws.Cells(j, 13).Value And ws.Cells(j + 1, 13).Value > stock_volume Then
stock_volume = ws.Cells(j + 1, 13).Value
ticker_val_stock = ws.Cells(j + 1, 10).Value
ElseIf ws.Cells(j, 13).Value > ws.Cells(j + 1, 13).Value And ws.Cells(j, 13).Value > stock_volume Then
stock_volume = ws.Cells(j, 13).Value
ticker_val_stock = ws.Cells(j, 10).Value
Else
stock_volume = stock_volume
ticker_val_stock = ticker_val_stock
End If

Next j

' Setting up the table
ws.Range("P2").Value = "Greatest % increase"
ws.Range("P3").Value = "Greatest % decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"
ws.Range("P2:P4").Font.Bold = True
ws.Range("Q1:R1").Font.Bold = True

'Putting the Ticker Values in the Table

ws.Range("Q2").Value = ticker_val_high
ws.Range("Q3").Value = ticker_val_low
ws.Range("Q4").Value = ticker_val_stock


'Putting the Values in the Table

ws.Range("R2").Value = greatest_percent
ws.Range("R3").Value = lowest_percent
ws.Range("R4").Value = stock_volume
 
 
ws.Range("L:L").NumberFormat = "0.00%"
ws.Range("R2:R3").NumberFormat = "0.00%"

Next ws


End Sub
