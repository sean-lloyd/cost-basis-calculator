Option Explicit

Sub Liquidate()
    Dim wsEvent As Worksheet, wsUTXO As Worksheet, wsDash As Worksheet
    Dim iLastRow_Evt As Long, iLastRow_Utxo As Long, iFirstRow As Long, rEvt As Long, rUtxo As Long
    Dim currYear As Integer

    Set wsEvent = Worksheets.Item("Events")
    Set wsUTXO = Worksheets.Item("UTXOs")
    Set wsDash = Worksheets.Item("Dashboard")
    currYear = wsDash.Range(CurrentYear).Value

    iFirstRow = 2
    iLastRow_Evt = wsEvent.Range("A1048576").End(xlUp).Row

    ' Loop through each taxable event
    Dim evtDate As Date, evtSymbol As String, evtAction As String, evtTXID As String
    Dim evtVolume As Double, evtPrice As Double
    For rEvt = iFirstRow To iLastRow_Evt
        ' Collect: Date, Symbol
        evtDate = wsEvent.Cells(rEvt, EVT_Date).Value
        evtAction = wsEvent.Cells(rEvt, EVT_Action).Value
        evtSymbol = wsEvent.Cells(rEvt, EVT_Symbol).Value
        evtVolume = wsEvent.Cells(rEvt, EVT_Volume).Value
        evtPrice = wsEvent.Cells(rEvt, EVT_PriceUSD).Value
        evtTXID = wsEvent.Cells(rEvt, EVT_TXID).Value

        If evtAction <> "INCOME" Then

            If FormatDateTime(evtDate, 3) = "12:00:00 AM" Then evtDate = ChangeTime_11_59(evtDate)  ' Chang time 11:59 PM when there is no time

            ' Loop through UTXOs to liquidate
            iLastRow_Utxo = wsUTXO.Range("A1048576").End(xlUp).Row
            Dim lUTXO_Symbol As String, l_Currency As String, lUTXO_DateAcquired As Date, lUTXO_DateSold As Date
            Dim lUTXO_PriceUSD As Double, lUTXO_PriceUSD_Evt As Double, lUTXO_VolumeOpen As Double, l_Proceeds As Double, l_CostBasis As Double, l_Gain As Double
            Dim lUTXO_TXID_Liquidated As String, lUTXO_Unmatched As String
            Dim liquidation As liquidation
            
            For rUtxo = iFirstRow To iLastRow_Utxo
                lUTXO_Symbol = wsUTXO.Cells(rUtxo, UTXO_Symbol).Value
                lUTXO_DateAcquired = wsUTXO.Cells(rUtxo, UTXO_DateAcquired).Value
                lUTXO_VolumeOpen = wsUTXO.Cells(rUtxo, UTXO_CY_CB_Vol_Open).Value
                lUTXO_PriceUSD = wsUTXO.Cells(rUtxo, UTXO_PriceUSD).Value

                If lUTXO_Symbol = evtSymbol And lUTXO_DateAcquired <= evtDate And lUTXO_VolumeOpen > 0 Then
                    Set liquidation = New liquidation

                    With liquidation
                        .year = currYear
                        .symbol = evtSymbol
                        .action = evtAction
                        .dateAcquired = lUTXO_DateAcquired
                        .dateSold = evtDate
                        .lcurrency = "USD"
                        .TXID = evtTXID
                        .lUTXO_TXID = wsUTXO.Cells(rUtxo, UTXO_TXID).Value
                        .unmatched = ""
                    End With

                    ' Calculate liquidation amounts
                    If evtVolume >= lUTXO_VolumeOpen Then ' Full liquidation
                        liquidation.volume = lUTXO_VolumeOpen
                    Else ' Partial liquidation
                        liquidation.volume = evtVolume
                    End If

                    liquidation.proceeds = (liquidation.volume * evtPrice)
                    liquidation.costBasis = (liquidation.volume * lUTXO_PriceUSD)
                    liquidation.gain = (liquidation.proceeds - liquidation.costBasis)
                    
                    ' Perform liquidation
                    ' Update Liquidation tab
                    liquidation.WriteOut

                    ' Update UTXO tab
                    wsUTXO.Cells(rUtxo, UTXO_CY_CB_Change).Value = liquidation.costBasis + wsUTXO.Cells(rUtxo, UTXO_CY_CB_Change).Value
                    wsUTXO.Cells(rUtxo, UTXO_CY_CB_Vol_Change).Value = liquidation.volume + wsUTXO.Cells(rUtxo, UTXO_CY_CB_Vol_Change).Value
                    wsUTXO.Cells(rUtxo, UTXO_CY_CB_Vol_Open).Value = lUTXO_VolumeOpen - liquidation.volume
                    wsUTXO.Cells(rUtxo, UTXO_LiqTXIDs).Value = wsUTXO.Cells(rUtxo, UTXO_LiqTXIDs).Value & ", " & evtTXID

                    evtVolume = evtVolume - liquidation.volume
                End If ' Main Liquidation

                If evtVolume = 0 Then Exit For

            Next rUtxo

            ' Unmatched: Events that do not have UTXOs available to liquidate
            If Round(evtVolume, 8) > 0 Then
                Set liquidation = New liquidation

                With liquidation
                    .year = currYear
                    .symbol = evtSymbol
                    .action = evtAction
                  ' .dateAcquired = Empty
                    .dateSold = evtDate
                    .lcurrency = "USD"
                    .TXID = evtTXID
                    .lUTXO_TXID = ""
                    .unmatched = "Y"
                    .volume = evtVolume
                    .proceeds = (liquidation.volume * evtPrice)
                    .costBasis = 0
                    .gain = liquidation.proceeds
                End With
                
                liquidation.WriteOut
            End If ' Unmatched

        End If ' <> Income

    Next rEvt

End Sub

Sub ClearLiquidations()
    Dim wsLiq As Worksheet, wsDash As Worksheet
    Dim iLastRow As Long, iFirstRow As Long, r As Long
    Dim iCurrYear As Integer
    
    Set wsLiq = Worksheets.Item("Liquidations")
    Set wsDash = Worksheets.Item("Dashboard")
    iFirstRow = 2
    iLastRow = wsLiq.Cells(Rows.Count, "A").End(xlUp).Row
    iCurrYear = wsDash.Range(CurrentYear).Value
    
    'Loop through each row in the sheet
    For r = iLastRow To 2 Step -1 'start from the last row and go backwards to avoid skipping rows after deletion
    
        'Check if the year in column A is 2023
        If wsLiq.Cells(r, LIQ_Year).Value = iCurrYear Then
        
            wsLiq.Rows(r).Delete
        
        End If
    
    Next r

End Sub

Private Function ChangeTime_11_59(dt_tm As Date)
    Dim dt As String, tm As String
    
    dt = FormatDateTime(dt_tm, 2)
    tm = FormatDateTime(dt_tm, 3)
    
    ChangeTime_11_59 = CDate(dt & " " & "11:59:59 PM")

End Function
