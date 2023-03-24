Sub PrintTXF()
    ' Declare Workbook variable
    Dim wsLiq As Worksheet, iFirstRow As Long, iLastRow As Long, r As Long
    Dim wsDash As Worksheet, CY As String

    ' Set worksheet variables
    Set wsLiq = Worksheets.Item("Liquidations")
    iFirstRow = 2
    iLastRow = wsLiq.Range("A1048576").End(xlUp).Row

    Set wsDash = Worksheets.Item("Dashboard")
    CY = wsDash.Range(CurrentYear).Value

    ' Decare file variables
    Dim FilePath As String
    Dim FileNum As Integer

    ' Set the file path and name
    FilePath = ThisWorkbook.Path & "\" & CY & ".txf"

    ' Get the next available file number
    FileNum = FreeFile()

    ' Create and open the file for writing
    Open FilePath For Output As #FileNum

    ' Write the header record to the file
    Print #FileNum, "V042"
    Print #FileNum, "ASeanTaxes"
    Print #FileNum, "D" & Format(Date, "mm/dd/yyyy")
    Print #FileNum, "^"

    ' Write the line records
    For r = iFirstRow To iLastRow
        ' Declare TXF Fields
        Dim symbol As String, vol As String
        Dim buyDate As Date, sellDate As Date
        Dim proceeds As Double, basis As Double
        Dim refLineNum As String

        With wsLiq
            symbol = .Cells(r, LIQ_Symbol).Value
            vol = .Cells(r, LIQ_Volume).Value
            buyDate = .Cells(r, LIQ_DateAcquired).Value
            sellDate = .Cells(r, LIQ_DateSold).Value
            proceeds = (.Cells(r, LIQ_Proceeds).Value)
            basis = Cells(r, LIQ_CostBasis).Value
        End With
        
        ' Short-Term = Line 712, Long-Term = Line 714 (Form 8949-C)
        refLineNum = IIf((sellDate - buyDate) < 365, "712", "714")
        
        ' Pring lines to file
        Print #FileNum, "TD"
        Print #FileNum, "N" & refLineNum
        Print #FileNum, "C1"
        Print #FileNum, "L1"
        Print #FileNum, "P" & vol & " " & symbol
        Print #FileNum, "D" & Format(buyDate, "mm/dd/yyyy")
        Print #FileNum, "D" & Format(sellDate, "mm/dd/yyyy")
        Print #FileNum, Format(Round(basis,2), "$#0.00")
        Print #FileNum, Format(Round(proceeds,2), "$#0.00")
        Print #FileNum, "^"
    Next r

    ' Close the file
    Close #FileNum

    ' Notify the user
    MsgBox "Tax file created successfully at: " & FilePath, vbInformation, "Success"
End Sub
