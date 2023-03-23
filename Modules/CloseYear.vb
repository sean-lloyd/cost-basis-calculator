Option Explicit

Sub CopyAndRenameSheet(sheetName, suffix)

    ' Declare variables
    Dim originalSheet As Worksheet
    Dim newSheet As Worksheet
    Dim sheetToDelete As Worksheet

    ' Check if the sheet named "UTXOs" exists
    Set originalSheet = ThisWorkbook.Worksheets(sheetName)

    ' Check if the sheet named "UTXOs BegBal" exists
    On Error Resume Next
    Set sheetToDelete = ThisWorkbook.Worksheets(sheetName & suffix)
    On Error GoTo 0

    ' Delete the "UTXOs BegBal" sheet if it exists
    If Not sheetToDelete Is Nothing Then
        Application.DisplayAlerts = False
        sheetToDelete.Delete
        Application.DisplayAlerts = True
    End If

    ' Make a copy of the sheet "UTXOs"
    originalSheet.Copy After:=originalSheet
    ' Set the new sheet as the active sheet
    Set newSheet = ActiveSheet
    ' Rename the new sheet
    newSheet.Name = sheetName & suffix

End Sub

