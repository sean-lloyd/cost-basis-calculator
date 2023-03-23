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
    wsDash.Range(CurrentYear).Value = CY + 1
    Call Utilities.CopyContentsBetweenSheets("UTXOs", "UTXOs_BegBal")

    ' Save with new name
    Call SaveAsNewYear(CY + 1)

    Application.ScreenUpdating = True

    ' Inform the user
    Dim msg As String
    msg = "Year " & CY & " is now closed and saved in a separate file." & vbNewLine & "You are now in a new workbook for Year " & (CY + 1) & "."
    MsgBox msg

End Sub

Private Sub SaveAsNewYear(newYear As String)
' Replace the last four characters of the old name with the new name and save the file
    Dim oldName As String
    Dim newName As String
    
    oldName = ActiveWorkbook.FullName
    newName = Left(oldName, Len(oldName) - 9) & newYear & ".xlsm"
    
    ActiveWorkbook.SaveAs newName
    
End Sub