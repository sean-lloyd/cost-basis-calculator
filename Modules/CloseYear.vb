Option Explicit

Sub CloseYear()
    Dim wsDash As Worksheet, CY As String

    Application.ScreenUpdating = False

    ' Get current year
    Set wsDash = Worksheets.Item("Dashboard")
    CY = wsDash.Range(CurrentYear).Value

    ' Save the current workbook
    ThisWorkbook.Save

    ' Back up the file
    Call Utilities.SaveBackupCopy("Backups")

    ' Create file for the new year
    Call Utilities.CopyContentsBetweenSheets("UTXOs", "UTXOs_BegBal") ' Set next year's beginning balance
    Call Create_UTXOs.ClearUTXO
    Call Create_Events.ClearEvents
    Call Liquidate_Events.ClearLiquidations

    wsDash.Activate
    wsDash.Range(CurrentYear).Value = CY + 1

    ' Save with new name
    Call SaveAsNewYear(CY + 1)

    Application.ScreenUpdating = True

    ' Inform the user
    Dim msg As String
    msg = "Year " & CY & " is now closed and saved in a separate file." & vbNewLine & "You are now in a new workbook for Year " & (CY + 1) & "."
    MsgBox msg

End Sub

Private Sub CloseUTXO()
    Dim wsUTXO As Worksheet
    Dim iFirstRow As Long, iLastRow As Long, r As Long
    Dim CB_Open As Double

    Set wsUTXO = Worksheets.Item("UTXOs")

    iFirstRow = 2
    iLastRow = wsUTXO.Range("A1048576").End(xlUp).Row
    
    ' Loop through the UTXOs and 
    For r = iLastRow To iFirstRow Step -1 'start from the last row and go backwards to avoid skipping rows after deletion
    
        CB_Open = wsUTXO.Cells(r, UTXO_CY_CB_Vol_Open).Value

        If CB_Open = 0 Then
            wsUTXO.Rows(r).Delete 'Delete when fully liquidated
        Else
            ' Update the long-term columns
            wsUTXO.Cells(r, UTXO_CostBasisVolumeOpen).Value = CB_Open
            wsUTXO.Cells(r, UTXO_CostBasisOpenUSD).Value = Round(CB_Open * wsUTXO.Cells(r, UTXO_PriceUSD).Value,2)
            
            ' Set the current-year columns = 0
            wsUTXO.Cells(r, UTXO_CY_CB_Change).Value = 0
            wsUTXO.Cells(r, UTXO_CY_CB_Vol_Change).Value = 0
            wsUTXO.Cells(r, UTXO_CY_CB_Vol_Open).Value = 0
        End IF

    Next r

End Sub

Private Sub SaveAsNewYear(newYear As String)
' Replace the last four characters of the old name with the new name and save the file
    Dim oldName As String
    Dim newName As String
    
    oldName = ActiveWorkbook.FullName
    newName = Left(oldName, Len(oldName) - 9) & newYear & ".xlsm"
    
    ActiveWorkbook.SaveAs newName
    
End Sub