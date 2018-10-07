Attribute VB_Name = "Module6"
Sub TickerAndVolumeSummary()

Dim Ticker As String
Dim TickerVolume As Double

TickerVolume = 0
  
Dim VolumeTotalRow As Integer

VolumeTotalRow = 2

Dim lastRow As Variant

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

    TickerVolume = TickerVolume + Cells(i, 7).Value

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    tickersymbol = Cells(i, 1).Value

    Cells(VolumeTotalRow, 9).Value = tickersymbol
    Cells(VolumeTotalRow, 10).Value = TickerVolume

    VolumeTotalRow = VolumeTotalRow + 1

    TickerVolume = 0

    End If

  Next i

End Sub
