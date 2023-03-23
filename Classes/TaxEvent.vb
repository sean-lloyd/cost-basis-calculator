Option Explicit

Public transDate As Date
Public action As String
Public symbol As String
Public volume As Double
Public priceUSD As Double
Public txid As String

Public Sub WriteOut()
    Dim wsEvent As Worksheet, r As Long
    Set wsEvent = Worksheets.Item("Events")
    r = wsEvent.Range("A1048576").End(xlUp).Row + 1

    With wsEvent
        .Cells(r, EVT_Date).Value = transDate
        .Cells(r, EVT_Action).Value = action
        .Cells(r, EVT_Symbol).Value = symbol
        .Cells(r, EVT_Volume).Value = volume
        .Cells(r, EVT_PriceUSD).Value = priceUSD
        .Cells(r, EVT_TXID).Value = txid
    End With

End Sub
