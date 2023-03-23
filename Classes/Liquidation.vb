Option Explicit

Public year as Integer
Public symbol As String
Public action As String
Public dateAcquired As Date
Public dateSold As Date
Public volume As Double
Public proceeds As Double
Public costBasis As Double
Public gain As Double
Public lcurrency As String
Public TXID As String
Public lUTXO_TXID As String
Public unmatched As String

Public Sub WriteOut()
    Dim wsLiq As Worksheet, r As Long
    Set wsLiq = Worksheets.Item("Liquidations")
    r = wsLiq.Range("A1048576").End(xlUp).Row + 1

    With wsLiq
        .Cells(r, LIQ_Year).Value = year
        .Cells(r, LIQ_Symbol).Value = symbol
        .Cells(r, LIQ_Action).Value = action
        .Cells(r, LIQ_DateAcquired).Value = dateAcquired
        .Cells(r, LIQ_DateSold).Value = dateSold
        .Cells(r, LIQ_Volume).Value = volume
        .Cells(r, LIQ_Proceeds).Value = proceeds
        .Cells(r, LIQ_CostBasis).Value = costBasis
        .Cells(r, LIQ_Gain).Value = gain
        .Cells(r, LIQ_Currency).Value = lcurrency
        .Cells(r, LIQ_TXID).Value = TXID
        .Cells(r, LIQ_UTXO_TXID).Value = lUTXO_TXID
        .Cells(r, LIQ_Unmatched).Value = unmatched
    End With

End Sub
