Option Explicit

Sub CreateEvents()
    Dim wsDash As Worksheet
    Dim currYear As Integer

    Set wsDash = Worksheets.Item("Dashboard")
    currYear = wsDash.Range(CurrentYear).Value

    Call TradingEvents(currYear)
    Call SpendingEvents(currYear)
    Call IncomeEvents(currYear)
    Call SortEvents

End Sub

Private Sub TradingEvents(currYear As Integer)
    Dim wsTR As Worksheet
    Dim taxEvent As taxEvent
    Dim iFirstRow As Long, iLastRow As Long, r As Long
    Dim tradeAction As String, tradeCurrency As String, tradeFeeCurrency As String, tradeSymbol As String, eligible As String
    Dim recordYear As Integer

    Set wsTR = Worksheets.Item("Trading")
    iFirstRow = 2
    iLastRow = wsTR.Range("A1048576").End(xlUp).Row
    eligible = "N"

    For r = iFirstRow To iLastRow
        Set taxEvent = New taxEvent
        
        ' Select record valid year
        recordYear = wsTR.Cells(r, TR_Year).Value

        If recordYear = currYear Then

            tradeSymbol = wsTR.Cells(r, TR_Symbol).Value
            tradeCurrency = wsTR.Cells(r, TR_Currency).Value
            tradeFeeCurrency = wsTR.Cells(r, TR_FeeCurrency).Value
            tradeAction = wsTR.Cells(r, TR_Action).Value

            If tradeAction = "SELL" Then eligible = "Y"
            If tradeAction = "BUY" And tradeCurrency <> "USD" Then eligible = "Y"

            If eligible = "Y" Then
                taxEvent.transDate = wsTR.Cells(r, TR_Date).Value
                
                If tradeAction = "SELL" Then
                    taxEvent.action = wsTR.Cells(r, TR_Action).Value
                    taxEvent.symbol = tradeSymbol
                    taxEvent.volume = wsTR.Cells(r, TR_Volume).Value
                Else ' "BUY"
                    taxEvent.action = "SELL"
                    taxEvent.symbol = tradeCurrency
                    taxEvent.volume = wsTR.Cells(r, TR_CostProceeds).Value
                End If

                taxEvent.priceUSD = wsTR.Cells(r, TR_USDTotalCost).Value / taxEvent.volume
                taxEvent.TXID = wsTR.Cells(r, TR_TXID).Value

                If taxEvent.TXID = "7Vzut+aLV09zwxLmmoGqB6Tvm7o" Then
                    Dim test as Integer
                    test = "WHY DID THIS GET SELECTED?"
                End If

                taxEvent.WriteOut

                ' Tax events for fees that are not absorbed by the buying or selling currency
                If tradeFeeCurrency <> tradeSymbol And tradeFeeCurrency <> tradeCurrency Then
                    taxEvent.action = "SPEND"
                    taxEvent.symbol = tradeFeeCurrency
                    taxEvent.volume = wsTR.Cells(r, TR_Fee).Value
                    taxEvent.priceUSD = wsTR.Cells(r, TR_USDFeeCost).Value / taxEvent.volume
                    taxEvent.TXID = wsTR.Cells(r, TR_TXID).Value
                    taxEvent.WriteOut
                End If ' Fee Tax Event

            End If 'Eligible

            eligible = "N"

        End If 'Record = Current Year
    Next r

End Sub

Private Sub SpendingEvents(currYear As Integer)
    Dim wsSP As Worksheet
    Dim taxEvent As taxEvent
    Dim iFirstRow As Long, iLastRow As Long, r As Long
    Dim recordYear As Integer

    Set wsSP = Worksheets.Item("Spending")
    iFirstRow = 2
    iLastRow = wsSP.Range("A1048576").End(xlUp).Row

    For r = iFirstRow To iLastRow
        Set taxEvent = New taxEvent
        
        ' Select record valid year
        recordYear = wsSP.Cells(r, TR_Year).Value

        If recordYear = currYear Then

            With taxEvent
                .transDate = wsSP.Cells(r, SP_Date).Value
                .action = wsSP.Cells(r, SP_Action).Value
                .symbol = wsSP.Cells(r, SP_Symbol).Value
                .volume = wsSP.Cells(r, SP_Volume).Value
                .priceUSD = wsSP.Cells(r, SP_USDTotalCost).Value / taxEvent.volume
                .TXID = wsSP.Cells(r, SP_TXID).Value
            End With

            taxEvent.WriteOut
        End If
    Next r
End Sub

Private Sub IncomeEvents(currYear As Integer)
    Dim wsINC As Worksheet
    Dim taxEvent As taxEvent
    Dim iFirstRow As Long, iLastRow As Long, r As Long
    Dim recordYear As Integer

    Set wsINC = Worksheets.Item("Income")
    iFirstRow = 2
    iLastRow = wsINC.Range("A1048576").End(xlUp).Row


    For r = iFirstRow To iLastRow
        Set taxEvent = New taxEvent
        
        ' Select record valid year
        recordYear = wsINC.Cells(r, TR_Year).Value

        If recordYear = currYear Then

            With taxEvent
                .transDate = wsINC.Cells(r, INC_Date).Value
                .action = wsINC.Cells(r, INC_Action).Value
                .symbol = wsINC.Cells(r, INC_Symbol).Value
                .volume = wsINC.Cells(r, INC_Volume).Value
                .priceUSD = wsINC.Cells(r, INC_USDTotalCost).Value / taxEvent.volume
                .TXID = wsINC.Cells(r, INC_TXID).Value
            End With

            taxEvent.WriteOut
        End If
    Next r
End Sub

Sub ClearEvents()
    Dim wsEvent As Worksheet, iLastRow As Long, iFirstRow As Long
    Set wsEvent = Worksheets.Item("Events")
    iFirstRow = 2
    iLastRow = wsEvent.Range("A1048576").End(xlUp).Row

    If iLastRow > iFirstRow Then
        wsEvent.Activate
        Range(EVT_Date & iFirstRow, EVT_TXID & iLastRow).Select
        Selection.ClearContents
        Range("A1").Select
    End If
End Sub

Private Sub SortEvents()
    Dim wsEvent As Worksheet
    Dim iLastRow  As Long, iFirstRow As Long
    Dim sortRange As String

    Set wsEvent = Worksheets.Item("Events")
    iFirstRow = 2
    iLastRow = wsEvent.Range("A1048576").End(xlUp).Row

    ' Sort by symbol
    sortRange = EVT_Date & "1:" & EVT_TXID & iLastRow
    wsEvent.Activate
    wsEvent.Range(sortRange).Sort _
        Key1:=Range(EVT_Date & "1"), Order1:=xlAscending, _
        Key2:=Range(EVT_Symbol & "1"), Order2:=xlAscending, _
        Key3:=Range(EVT_Volume & "1"), Order3:=xlDescending, _
        Header:=xlYes

End Sub