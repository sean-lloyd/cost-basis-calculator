Option Explicit

Sub SaveBackupCopy(Optional folderName As String)

    Dim originalFilePath As String
    Dim backupFilePath As String
    Dim backupFileName As String
    Dim fileDirectory As String
    Dim timeStamp As String

    ' Get the original file path and the directory where the file is saved
    originalFilePath = ThisWorkbook.FullName
    fileDirectory = Left(originalFilePath, InStrRev(originalFilePath, "\"))

    ' Add folder to directory, if provided
    If Len(folderName) > 0 Then fileDirectory = fileDirectory & folderName & "\"
    
    ' Create a date-time stamp in format yyyy-mm-dd_hh-mm-ss
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    
    ' Create the backup file name using the original file name, "backup", and the date-time stamp
    backupFileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_bkup_" & timeStamp & ".xlsm"
    
    ' Combine the file directory and backup file name
    backupFilePath = fileDirectory & backupFileName
    
    ' Save the backup copy in the same directory as the working file
    ThisWorkbook.SaveCopyAs backupFilePath

End Sub

Sub CopyAndRenameSheet(sheetName, suffix)

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

Sub CopyContentsBetweenSheets(fromSheet As String, toSheet As String)

    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Worksheets(fromSheet)
    Set destinationSheet = ThisWorkbook.Worksheets(toSheet)

    ' Clear the contents of the destination sheet "UTXOs"
    destinationSheet.Cells.ClearContents

    ' Copy the contents from the source sheet to the destination sheet 
    sourceSheet.UsedRange.Copy Destination:=destinationSheet.Cells(1, 1)

End Sub