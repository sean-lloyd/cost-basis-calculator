Option Explicit

Sub CreateUTXOs()
    Dim wsDash As Worksheet, wsTR As Worksheet, wsSP As Worksheet, wsINC As Worksheet
    Dim iLastRow_TR As Long, iLastRow_SP As Long, iLastRow_INC As Long
    Dim wsUTXO As Worksheet, iLastRow_UT As Long
    Dim iFirstRow As Long, r As Long
    Dim iCurrYear As Integer, recordYear As Integer
    Dim tradeAction As String, tradeCurrency As String, tradeFeeCurrency As String, eligible As String
    Dim utxo As utxo
    
    ' Set worksheet references
    Set wsDash = Worksheets.Item("Dashboard")
    Set wsTR = Worksheets.Item("Trading")
    Set wsSP = Worksheets.Item("Spending")
    Set wsINC = Worksheets.Item("Income")
    Set wsUTXO = Worksheets.Item("UTXOs")

    ' Set starting variables
    iCurrYear = wsDash.Range(CurrentYear).Value
    iFirstRow = 2
    iLastRow_TR = wsTR.Range("A1048576").End(xlUp).Row
    iLastRow_SP = wsSP.Range("A1048576").End(xlUp).Row
    iLastRow_INC = wsINC.Range("A1048576").End(xlUp).Row
    eligible = "N"

    ' Loop Through Trading
    For r = iFirstRow To iLastRow_TR
        Set utxo = New utxo
        
        ' Select record valid year and open CB
        recordYear = wsTR.Cells(r, TR_Year).Value

        If recordYear = iCurrYear Then

            tradeAction = wsTR.Cells(r, TR_Action).Value
            tradeCurrency = wsTR.Cells(r, TR_Currency).Value
            tradeFeeCurrency = wsTR.Cells(r, TR_FeeCurrency).Value

            ' Set eligbility
            If tradeAction = "BUY" Then eligible = "Y"
            If tradeAction = "SELL" And tradeCurrency <> "USD" Then eligible = "Y"

            If eligible = "Y" Then
                If tradeAction = "BUY" Then
                    utxo.symbol = wsTR.Cells(r, TR_Symbol).Value
                    utxo.volume = wsTR.Cells(r, TR_Volume).Value

                    ' Subtract the trade fee when it was paid out of the symbol purchased
                    If utxo.symbol = tradeFeeCurrency Then
                        utxo.volume = (utxo.volume - wsTR.Cells(r, TR_Fee).Value)
                    End If

                Else ' "SELL"
                    utxo.symbol = tradeCurrency
                    utxo.volume = wsTR.Cells(r, TR_CostProceeds).Value
                End If
                
                utxo.year = wsTR.Cells(r, TR_Year).Value
                utxo.dateAcquired = wsTR.Cells(r, TR_Date).Value
                utxo.price = wsTR.Cells(r, TR_USDTotalCost).Value / utxo.volume
                utxo.costBasis = utxo.volume * utxo.price
                utxo.costBasisOpen = utxo.costBasis
                utxo.costBasisVolOpen = utxo.volume
                utxo.category = "Trading"
                utxo.TXID = wsTR.Cells(r, TR_TXID).Value

                utxo.WriteOut

            End If 'Eligible

            eligible = "N"

        End If 'Current Year

    Next r

    'Loop Through Income
    For r = iFirstRow To iLastRow_INC
        Set utxo = New utxo

        ' Select record valid year and open CB
        recordYear = wsINC.Cells(r, INC_Year).Value

        If recordYear = iCurrYear Then
        
            With utxo
                .symbol = wsINC.Cells(r, INC_Symbol).Value
                .year = wsINC.Cells(r, INC_Year).Value
                .dateAcquired = wsINC.Cells(r, INC_Date).Value
                .volume = wsINC.Cells(r, INC_Volume).Value
                .price = wsINC.Cells(r, INC_USDTotalCost).Value / utxo.volume
                .costBasis = utxo.volume * utxo.price
                .costBasisOpen = utxo.costBasis
                .costBasisVolOpen = utxo.volume
                .category = "Income"
                .TXID = wsINC.Cells(r, INC_TXID).Value
            End With
            
            utxo.WriteOut
        End If

    Next r
    
    Call SortUTXO
    
End Sub

Sub ClearUTXO()
    Dim wsUTXO As Worksheet, wsDash As Worksheet
    Dim iLastRow As Long, iFirstRow As Long, r As Long
    Dim iCurrYear As Integer

    Set wsUTXO = Worksheets.Item("UTXOs")
    Set wsDash = Worksheets.Item("Dashboard")
    iFirstRow = 2
    iLastRow = wsUTXO.Cells(Rows.Count, "A").End(xlUp).Row
    iCurrYear = wsDash.Range(CurrentYear).Value
    
    'Loop through each row in the sheet
    For r = iLastRow To 2 Step -1 'start from the last row and go backwards to avoid skipping rows after deletion
    
        'Check if the year in column A is 2023
        If wsUTXO.Cells(r, UTXO_Year).Value = iCurrYear Then
        
            wsUTXO.Rows(r).Delete
        
        End If
    
    Next r
    
End Sub

Private Sub SortUTXO()
    Dim wsUTXO, wsDash As Worksheet
    Dim iLastRow, iFirstRow As Long, iSymStart As Long, iSymEnd As Long, r As Long
    Dim sGL_Method As String, sSym_Curr As String, sSym_Prev As String, sSym_Next As String
    Dim r1 As String, r2 As String, sortRange As String
    Dim iCurrYear As Integer
    
    Set wsUTXO = Worksheets.Item("UTXOs")
    Set wsDash = Worksheets.Item("Dashboard")
    iFirstRow = 2
    iLastRow = wsUTXO.Range("A1048576").End(xlUp).Row
    iSymStart = 0
    iSymEnd = 0
    iCurrYear = wsDash.Range(CurrentYear).Value

    ' Sort by symbol
    sortRange = UTXO_Symbol & "1:" & UTXO_LiqTXIDs & iLastRow
    wsUTXO.Activate
    wsUTXO.Range(sortRange).Sort _
        Key1:=Range(UTXO_Symbol & "1"), Order1:=xlAscending, Header:=xlYes

    ' Sort each symbol's range
    For r = iFirstRow To iLastRow
        sSym_Curr = wsUTXO.Range(UTXO_Symbol & r).Value
        sSym_Prev = wsUTXO.Range(UTXO_Symbol & (r - 1)).Value
        sSym_Next = wsUTXO.Range(UTXO_Symbol & (r + 1)).Value

        ' Set Start Row for Symbol's Range
        If sSym_Curr <> sSym_Prev Then iSymStart = r

        ' Set End Row for Symbol's Range
        If sSym_Curr <> sSym_Next Then iSymEnd = r

        If iSymEnd > 0 Then
            ' Get range
            r1 = UTXO_Symbol & iSymStart
            r2 = UTXO_LiqTXIDs & iSymEnd
            
            ' Get Sort Method
            sGL_Method = GetSortMethod(sSym_Curr)

            ' Apply sort
                Call ApplySortMethod(r1, r2, iSymStart, sGL_Method)

            iSymEnd = 0
        End If
    Next r

    Range("A1").Select
End Sub

Private Function GetSortMethod(symbol As String) As String
    Dim wsDash As Worksheet
    Dim r As Long, iLastRow As Long
    Dim sCurrSymbol As String
    Set wsDash = Worksheets.Item("Dashboard")
    iLastRow = wsDash.Range(GL_Symbol_C & "1048576").End(xlUp).Row
    GetSortMethod = "None"

    ' Lookup symbol on Dashboard
    For r = GL_Method_FirstRow To iLastRow
        sCurrSymbol = wsDash.Range(GL_Symbol_C & r).Value

        If sCurrSymbol = symbol Then
            GetSortMethod = wsDash.Range(GL_Method_C & r).Value
            Exit For
        End If
    Next r
    
    If GetSortMethod = "None" Then GetSortMethod = "FIFO"
    
End Function

Private Sub ApplySortMethod(a As String, b As String, startRow As Long, method As String)
    Dim wsUTXO As Worksheet

    Set wsUTXO = Worksheets.Item("UTXOs")
    wsUTXO.Activate
    Range(a, b).Select

    Select Case method

        Case "FIFO"

            wsUTXO.Range(a, b).Sort _
                Key1:=Range(UTXO_DateAcquired & startRow), Order1:=xlAscending, _
                Key2:=Range(UTXO_CostBasisVolumeOpen & startRow), Order2:=xlAscending, _
                Header:=xlNo

        Case "LIFO"

            wsUTXO.Range(a, b).Sort _
                Key1:=Range(UTXO_DateAcquired & startRow), Order1:=xlDescending, _
                Key2:=Range(UTXO_CostBasisVolumeOpen & startRow), Order2:=xlAscending, _
                Header:=xlNo

        Case "HCFO"

            wsUTXO.Range(a, b).Sort _
                Key1:=Range(UTXO_CostBasisUSD & startRow), Order1:=xlDescending, _
                Key2:=Range(UTXO_CostBasisVolumeOpen & startRow), Order2:=xlAscending, _
                Header:=xlNo

        Case "LCFO"

            wsUTXO.Range(a, b).Sort _
                Key1:=Range(UTXO_CostBasisUSD & startRow), Order1:=xlAscending, _
                Key2:=Range(UTXO_CostBasisVolumeOpen & startRow), Order2:=xlAscending, _
                Header:=xlNo

        Case "HPFO"

            wsUTXO.Range(a, b).Sort _
                Key1:=Range(UTXO_PriceUSD & startRow), Order1:=xlDescending, _
                Key2:=Range(UTXO_CostBasisVolumeOpen & startRow), Order2:=xlAscending, _
                Header:=xlNo

        Case "LPFO"

            wsUTXO.Range(a, b).Sort _
                Key1:=Range(UTXO_PriceUSD & startRow), Order1:=xlAscending, _
                Key2:=Range(UTXO_CostBasisVolumeOpen & startRow), Order2:=xlAscending, _
                Header:=xlNo

        Case Else 'Default = FIFO
        
            wsUTXO.Range(a, b).Sort _
                Key1:=Range(UTXO_DateAcquired & startRow), Order1:=xlAscending, _
                Key2:=Range(UTXO_CostBasisVolumeOpen & startRow), Order2:=xlAscending, _
                Header:=xlNo
    
    End Select

End Sub
