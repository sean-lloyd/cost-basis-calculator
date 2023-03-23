Option Explicit

Public symbol As String
Public year As Integer
Public dateAcquired As Date
Public volume As Double
Public price As Double
Public costBasis As Double
Public costBasisOpen As Double
Public costBasisVolOpen As Double
Public category As String
Public txid As String
Public CY_CB_USD As Double
Public CY_CB_Vol_Change As Double
Public CY_CB_Vol_Open As Double
Public LiqTXIDs As String

Public Sub Liquidate(r)
    Dim wsUTXO As Worksheet
    Set wsUTXO = Worksheets.Item("UTXOs")
    wsUTXO.Cells(r, UTXO_CY_CB_USD_Open).Value = CY_CB_USD
    wsUTXO.Cells(r, UTXO_CY_CB_Vol_Change).Value = CY_CB_Vol_Change
    wsUTXO.Cells(r, UTXO_CY_CB_Vol_Open).Value = CY_CB_Vol_Open
    wsUTXO.Cells(r, UTXO_LiqTXIDs).Value = wsUTXO.Cells(r, UTXO_LiqTXIDs).Value & "," & LiqTXIDs
End Sub

Public Sub WriteOut()
    Dim wsUTXO As Worksheet, r As Long
    Set wsUTXO = Worksheets.Item("UTXOs")
    r = wsUTXO.Range("A1048576").End(xlUp).Row + 1

    With wsUTXO
        .Cells(r, UTXO_Symbol).Value = symbol
        .Cells(r, UTXO_Year).Value = year
        .Cells(r, UTXO_DateAcquired).Value = dateAcquired
        .Cells(r, UTXO_Volume).Value = volume
        .Cells(r, UTXO_PriceUSD).Value = price
        .Cells(r, UTXO_CostBasisUSD).Value = costBasis
        .Cells(r, UTXO_CostBasisOpenUSD).Value = costBasisOpen
        .Cells(r, UTXO_CostBasisVolumeOpen).Value = costBasisVolOpen
        .Cells(r, UTXO_Category).Value = category
        .Cells(r, UTXO_TXID).Value = txid
        .Cells(r, UTXO_CY_CB_USD_Open).Value = price * costBasisVolOpen
        .Cells(r, UTXO_CY_CB_Vol_Change).Value = 0
        .Cells(r, UTXO_CY_CB_Vol_Open).Value = costBasisVolOpen
        .Cells(r, UTXO_LiqTXIDs).Value = ""
    End With

End Sub

