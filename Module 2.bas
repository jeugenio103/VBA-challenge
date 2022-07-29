Attribute VB_Name = "Module1"
Sub ticker()

'Labeling Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Last Row Calculator
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Ticker Name
Dim Ticker_Name As String

'Ticker Information
Dim Ticker_Opening As Double
Dim Ticker_Closing As Double
Dim Ticker_Volume As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

Ticker_Opening = 0
Ticker_Closing = 0
Ticker_Volume = 0
Yearly_Change = 0
Percent_Change = 0

'Keep Tracking of Location
Dim Table_Row As Integer
Table_Row = 2

'Loop for Ticker
For i = 2 To lastrow


If Ticker_Opening = 0 Then
    Ticker_Opening = Cells(i, 3)
End If

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker_Name = Cells(i, 1).Value
    Ticker_Volume = Ticker_Volume + Cells(i, 7)
    Ticker_Closing = Cells(i, 6)
    Yearly_Change = Ticker_Closing - Ticker_Opening
    Percent_Change = Yearly_Change / Ticker_Opening
    
'Printing in Summary Table
    Range("I" & Table_Row).Value = Ticker_Name
    Range("J" & Table_Row).Value = Yearly_Change
    Range("K" & Table_Row).Value = Format(Percent_Change, "Percent")
    Range("L" & Table_Row).Value = Ticker_Volume
If Percent_Change > 0 Then
    Cells(Table_Row, 10).Interior.ColorIndex = 4
Else
    Cells(Table_Row, 10).Interior.ColorIndex = 3
End If

'Reseting volume and adding summary table row
    Table_Row = Table_Row + 1
    Ticker_Volume = 0
    Ticker_Opening = 0
    
Else:
    Ticker_Volume = Ticker_Volume + Cells(i, 7)

End If

Next i



End Sub
