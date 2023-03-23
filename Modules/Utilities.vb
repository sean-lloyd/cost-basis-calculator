Option Explicit

Sub SaveBackupCopy()

    ' Declare variables
    Dim originalFilePath As String
    Dim backupFilePath As String
    Dim backupFileName As String
    Dim fileDirectory As String
    Dim timeStamp As String
    
    ' Get the original file path and the directory where the file is saved
    originalFilePath = ThisWorkbook.FullName
    fileDirectory = Left(originalFilePath, InStrRev(originalFilePath, "\"))
    
    ' Create a date-time stamp in format yyyy-mm-dd_hh-mm-ss
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    
    ' Create the backup file name using the original file name, "backup", and the date-time stamp
    backupFileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_bkup_" & timeStamp & ".xlsm"
    
    ' Combine the file directory and backup file name
    backupFilePath = fileDirectory & backupFileName
    
    ' Save the backup copy in the same directory as the working file
    ThisWorkbook.SaveCopyAs backupFilePath

End Sub
