Attribute VB_Name = "Stock"
'Apply other methods to all worksheets
Sub allSheets()

For Each ws In Worksheets
    ws.Select
    Call stockVolume
Next

End Sub


Sub stockVolume()

'Labels for tickers and volumes in J & K columns
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Total Stock Volume"

'Make the stock volume column bigger
Columns("J").ColumnWidth = 20

'i is a counter variable
Dim i As Long

Dim totalRows As Long
Dim ticker As String
Dim currentStockVolume As Double
Dim finalStockVolume As Double

Dim t As Integer

t = 1
finalStockVolume = 0

totalRows = Cells(rows.Count, 1).End(xlUp).Row

'loop to get total stock volume for each stock
For i = 2 To totalRows

    If Cells(i, "A").Value = Cells(i + 1, "A").Value Then
        currentStockVolume = Cells(i, "G").Value
        finalStockVolume = finalStockVolume + currentStockVolume

    Else
        ticker = Cells(i, "A").Value

        currentStockVolume = Cells(i, "G").Value
        finalStockVolume = finalStockVolume + currentStockVolume

        t = t + 1

        Cells(t, "I").Value = ticker
        Cells(t, "J").Value = finalStockVolume

        currentStockVolume = 0
        finalStockVolume = 0

      End If
        
Next i

 
End Sub




