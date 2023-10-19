# vba-challenge-new-
Sub stock()


Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate


Dim Ticker As String
Ticker = ""

Dim YearlyChange As Double
YearlyChange = 0

Dim InitialValue As Double
InitialValue = ws.Cells(2, 3).Value


Dim FinalValue As Double
FinalValue = 0

Dim PercentChange As Double
PercentChange = 0

Dim TotalVolume As Double
TotalVolume = 0

Dim SummaryRow As Integer
SummaryRow = 2

Dim GreatIncr As Double
GreatIncr = ws.Cells(2, 11).Value
Dim TickerIncr As String
TickerIncr = ""

Dim GreatDecr As Double
GreatDecr = ws.Cells(2, 11).Value
Dim TickerDecr As String
TickerDecr = ""

Dim GreatVol As Double
GreatVol = ws.Cells(2, 12).Value
Dim TickerVol As String
TickerVol = ""

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yealy Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


For i = 2 To Rows.Count
If IsEmpty(ws.Cells(i, 1)) Then
End If



If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value

ws.Range("I" & SummaryRow).Value = Ticker

FinalValue = Cells(i, 6).Value
YearlyChange = FinalValue - InitialValue
ws.Range("J" & SummaryRow).Value = YearlyChange

If YearlyChange < 0 Then
ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
ElseIf YearlyChange > 0 Then
ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
End If


PercentChange = (YearlyChange / InitialValue)
ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
ws.Range("K" & SummaryRow).Value = PercentChange


TotalVolume = TotalVolume + ws.Cells(i, 7).Value
ws.Range("L" & SummaryRow).Value = TotalVolume

SummaryRow = SummaryRow + 1
TotalVolume = 0
InitialValue = Cells(i + 1, 3).Value


Else
TotalVolume = TotalVolume + ws.Cells(i, 7).Value
End If


ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


If ws.Cells(i + 1, 11).Value < GreatIncr Then
GreatIncr = ws.Cells(i + 1, 11).Value
TickerIncr = ws.Cells(i + 1, 9).Value
End If


If ws.Cells(i + 1, 11).Value > GreatDecr Then
GreatDecr = ws.Cells(i + 1, 11).Value
TickerDecr = ws.Cells(i + 1, 9).Value
End If

If ws.Cells(i + 1, 12).Value > GreatVol Then
GreatVol = ws.Cells(i + 1, 12).Value
TickerVol = ws.Cells(i + 1, 9).Value
End If


ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q2").Value = GreatIncr
ws.Range("P2").Value = TickerIncr

ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q3").Value = GreatDecr
ws.Range("P3").Value = TickerDecr

ws.Range("Q4").Value = GreatVol
ws.Range("P4").Value = TickerVol


Next i

Next ws

End Sub
