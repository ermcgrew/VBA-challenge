Option Explicit
Sub yearly_change()
       
    Dim end_data As Long
    end_data = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim output_row As Integer
    output_row = 2
    
    Dim total_volume As LongLong
    Dim last_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    Dim first_open As Double
    first_open = Range("C2")
    
    'i = iterator for rows of data
    Dim i As Long
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
 
    Columns("I:L").AutoFit
    
    For i = 2 To end_data
        'check if ticker is the same as next, if different, compile info for output and reset for next ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'ticker name to output
            Cells(output_row, 9) = Cells(i, 1).Value
            
            'add current row to total volume
            total_volume = total_volume + Cells(i, 7).Value
            
            'print total volume in output
            Cells(output_row, 12) = total_volume
            
            'get last_close number
            last_close = Cells(i, 6).Value
            
            'yearly change
            yearly_change = last_close - first_open
            Cells(output_row, 10).Value = yearly_change
            
            'conditional formatting on yearly_change
                If yearly_change > 0 Then
                    Cells(output_row, 10).Interior.ColorIndex = 4
                Else
                    Cells(output_row, 10).Interior.ColorIndex = 3
                End If
            
            'percent change
                If first_open = 0 Then
                    percent_change = 0
                Else
                    percent_change = (yearly_change / first_open) * 100
                End If
            
            percent_change = WorksheetFunction.Round(percent_change, 2)
            Cells(output_row, 11).Value = Str(percent_change) & "%"
            
            'reset total_volume, last_close, yearly_change, percent_change for next ticker
            total_volume = 0
            last_close = 0
            yearly_change = 0
            percent_change = 0
            
            'store first open for next ticker for next output
                If Cells(i + 1, 3).Value = 0 Then
                'find next cell within that ticker with non-zero number
                    Dim j As Long 'iterator
                    For j = i + 1 To 5000
                        If Cells(j, 3).Value = 0 Then
                            j = j + 1
                        Else
                            first_open = Cells(j, 3).Value
                        End If
                    Next j
                
                Else
                    first_open = Cells(i + 1, 3).Value
            
                End If
            
            'move to next line in output for next ticker
            output_row = output_row + 1
            
        'if same, add volume from current row to total volume
        Else
            total_volume = total_volume + Cells(i, 7).Value
              
        End If
    
    Next i
End Sub


