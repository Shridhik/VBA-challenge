Attribute VB_Name = "Module1"
Sub LoopThroughWorksheets()
Dim ws As Worksheet
Dim CurrentTicker As String
Dim CurrentTickerOpen As Double
Dim CurrentTickerClose As Double
Dim YearlyChange As Double
Dim TotalVol As Double
Dim GreatestDecreaseTicker As String
Dim GreatestDecreaseValue As Double
Dim GreatestIncreaseTicker As String
Dim GreatestIncreaseValue As Double
Dim i, j As Integer
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
TotalVol = 0
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
OpenValue = Cells(2, 3).Value
CurrentTicker = Cells(2, 1).Value
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Geatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
j = 2
For i = 2 To lastRow 'goes through all rows'
      TotalVol = TotalVol + Cells(i, 7).Value 'Keeps adding onto Total Volume'
      CurrentTicker = Cells(i, 1).Value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'If the next ticker is not equal to the current ticker then'
           CurrentTickerClose = Cells(i, 6).Value 'Store Close Value'
           Cells(j, 9).Value = CurrentTicker  'Prints next Ticker'
           Cells(j, 10).Value = CurrentTickerClose - OpenValue  'Prints the yearly change' '
                 If Cells(j, 10).Value < 0 Then
                    Cells(j, 10).Interior.Color = vbRed
                Else
                    Cells(j, 10).Interior.Color = vbGreen
                End If
    Cells(j, 11).Value = (((CurrentTickerClose - OpenValue) / (OpenValue)) * 100)  'Prints the percent change'
                If Cells(2, 17).Value < Cells(j, 11).Value Then
                    Cells(2, 17).Value = Cells(j, 11).Value
                    Cells(2, 16).Value = Cells(j, 9).Value
                End If
                If Cells(3, 17).Value > Cells(j, 11).Value Then
                    Cells(3, 17).Value = Cells(j, 11).Value
                    Cells(3, 16).Value = Cells(j, 9).Value
                End If
           'Cells(2, 16).Value = Max(Cells(2, 16), Cells(j, 11).Value)'
   Cells(j, 12).Value = TotalVol  'Prints the total volume' '
                If Cells(4, 17).Value < Cells(j, 12).Value Then
                    Cells(4, 17).Value = Cells(j, 12).Value
                    Cells(4, 16).Value = Cells(j, 9).Value
                End If
           TotalVol = 0
           j = j + 1
          OpenValue = Cells(i + 1, 3).Value 'Store Open value'
    Else
    End If
Next i
    Next ws
End Sub
