Sub Final_test()

' Defining The Different Variables
Dim Opening_price As Double
Dim Closing_price As Double
Dim Difference As Double
Dim Percent_change As Double
Dim cumulative_vol As Double
Dim current_vol As Double



' Inserting the headers of the summary table
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
Range("J1:M1").Font.Bold = True


Let i = 2
For r = 1 To 797711

' Obtaining the Ticker and the opening price every time there is a new ticker stock
If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
Cells(i, 10).Value = Cells(r + 1, 1).Value
End If

' Tallying up the volume
current_vol = Cells(r + 1, 7).Value
cumulative_vol = cumulative_vol + current_vol

'Obtaining the opening Price for Stocks
If Cells(r + 1, 1).Value <> Cells(r, 1).Value And Cells(r + 1, 3).Value <> 0 Then
Opening_price = Cells(r + 1, 3)
' Adjusting the opening price if it starts of as a 0 and eventually becomes a different opening price on a different day
ElseIf Cells(r + 1, 1).Value = Cells(r, 1).Value And Cells(r, 3).Value = 0 And Cells(r + 1, 3).Value <> 0 Then
Opening_price = Cells(r + 1, 3).Value
' Opening price if it is  0 the whole way through
ElseIf Cells(r + 1, 1).Value = Cells(r, 1).Value And Cells(r, 3).Value = 0 And Cells(r + 1, 3).Value = 0 Then
Opening_price = 0
End If

'Obtaining the Closing price for Stocks
If Cells(r + 2, 1).Value <> Cells(r + 1, 1).Value Then
Closing_price = Cells(r + 1, 6)
'Avoiding the Dividing by 0 problem
If Opening_price = 0 And Closing_price = 0 Then
Percent_change = 0
Difference = 0
Else
Difference = Closing_price - Opening_price
Percent_change = Difference / Opening_price
End If
Cells(i, 12).Value = Percent_change
' Conditional Formatting for the Yearly column
Cells(i, 11).Value = Difference
If Difference < 0 Then
Cells(i, 11).Interior.ColorIndex = 3
Else
Cells(i, 11).Interior.ColorIndex = 4
End If

Cells(i, 13).Value = cumulative_vol

cumulative_vol = 0

i = i + 1

End If

Next r
 
' Formatting the Percent change row to percentage
Range("L:L").NumberFormat = "0.00%"

End Sub
