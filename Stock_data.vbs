
' This script is for the testing part of the census data
Sub stockdata():
'Creating the columns required for the calculations
    For Each ws In Worksheets
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
            
        'Declare in memory the variables required for the calculations
        Dim i, OpeningRow, LastRow, CurrentRow, YearlyChangeLastRow, PercentChangeLastRow As Long
        Dim Total, Change, PercentChange, OpeningPrice As Double
        Dim GetTickerName As String
        
        'Get the last row using 1st column
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Initialize theopening row to be used later and current row will point to the ticker being used in the active cell for calculations
        OpeningRow = 2
        CurrentRow = 2
        Change = 0
        ' Get the value for the 1st ticker .This will be updated in the loop while calculating
        OpeningPrice = ws.Cells(OpeningRow, 3).Value
            For i = 2 To LastRow
                    'Compare the ticker name values
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                       GetTickerName = ws.Cells(i, 1)                           ' update the ticker name
                       ws.Range("I" & CurrentRow).Value = GetTickerName
                       Change = (ws.Cells(i, 6).Value - OpeningPrice) 'Calculate the yearly chnage by subtracting the closed and open value
                       ws.Range("J" & CurrentRow).Value = Change ' Update the Value in column J for the yearly change
                       PercentChange = (Change / OpeningPrice) ' Calcute the percent change
                       
                       ws.Range("K" & CurrentRow).Value = PercentChange ' update it
                       ws.Range("K" & CurrentRow).Style = "Percent"
                       Total = Total + ws.Cells(i, 7).Value             ' get the total for the same ticker
                       ws.Range("L" & CurrentRow).Value = Total         'update in the column for the total
                       CurrentRow = CurrentRow + 1                      ' update the current row to the next one
                       Change = 0                                ' reset the yearly change for the next ticker
                       Total = 0                                        'reset the total value
                       OpeningPrice = ws.Cells(i + 1, 3).Value    ' get the value for the next ticker
                    Else
                       Total = Total + ws.Cells(i, 7).Value ' otherwise keep adding the total for the same ticker if the tivcker name matches
                    End If
                    
                   
            Next i
            
             ' Conditional formatting
            YearlyChangeLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row ' Get the last row for the year change
                For i = 2 To YearlyChangeLastRow '
                    If ws.Cells(i, 10).Value >= 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4 ' for positive change the color to green
                    Else
                        ws.Cells(i, 10).Interior.ColorIndex = 3 ' for the negative chnange to red
                    End If
                Next i
                'min and max percent chnage
            PercentChangeLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
             For i = 2 To PercentChangeLastRow
            
        'compare the values in thr column for the tickers to get the min and max values and corresponig ticker
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then ' use percent chnage column for comparison
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
    
                 ElseIf ws.Range("K" & i).Value < ws.Range("Q3").Value Then ' for min value
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                        
                End If
        ' find  total volume max and use column for the total stock for comparison
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            'Chnage the style
            ws.Range("Q2").Style = "Percent"
            ws.Range("Q3").Style = "Percent"
            'autofit the whole sheet columns and rows
            ws.Cells.EntireColumn.AutoFit
            ws.Cells.EntireRow.AutoFit
            ' move to thr next sheet

    Next ws
End Sub

