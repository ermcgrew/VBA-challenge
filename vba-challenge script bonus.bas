Sub greatest()

'headers
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Dim end_data As Long
end_data = Cells(Rows.Count, 1).End(xlUp).Row

'greatest % increase
Dim j As Long
Dim biggest_increase As Double
biggest_increase = 0#
Dim ticker_inc As String

    For j = 2 To end_data
        If Cells(j, 11).Value > biggest_increase Then
            biggest_increase = Cells(j, 11).Value
            ticker_inc = Cells(j, 9).Value
        End If
    Next j
    
Cells(2, 17).Value = FormatPercent(biggest_increase, 2)
Cells(2, 16).Value = ticker_inc


'greatest % decrease
Dim k As Long
Dim biggest_decrease As Double
biggest_decrease = 0#
Dim ticker_dec As String

    For k = 2 To end_data
        If Cells(k, 11).Value < biggest_decrease Then
            biggest_decrease = Cells(k, 11).Value
            ticker_dec = Cells(k, 9).Value
        End If
    Next k
    
Cells(3, 17).Value = FormatPercent(biggest_decrease, 2)
Cells(3, 16).Value = ticker_dec

'Total stock volume
Dim i As Long
Dim biggest_volume As LongLong
Dim ticker As String

    For i = 2 To end_data
        If Cells(i, 12).Value > biggest_volume Then
            biggest_volume = Cells(i, 12).Value
            ticker = Cells(i, 9).Value
        End If
    Next i
    
Cells(4, 17).Value = biggest_volume
Cells(4, 16).Value = ticker

Columns("O:Q").AutoFit

End Sub