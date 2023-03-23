Option Explicit

Sub RunAll()
    Dim wsDash As Worksheet, runThruCY As String, CY As String, year As Integer

    Application.ScreenUpdating = False
    Set wsDash = Worksheets.Item("Dashboard")
    CY = wsDash.Range(CurrentYear).Value

    Application.StatusBar = "Backing Up File..."
    Call Utilities.SaveBackupCopy

    Application.StatusBar = "Cleaning Up..."
    Call Create_UTXOs.ClearUTXO
    Call Create_Events.ClearEvents
    Call Liquidate_Events.ClearLiquidations

    ' Run the UTXOs, events, and liquidations
    Call RunYear(CY)

    wsDash.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
End Sub

Private Sub RunYear(year)
    Application.StatusBar = "Creating UTXOs..." & year
    Call Create_UTXOs.CreateUTXOs
    Application.StatusBar = "Creating Tax Events..." & year
    Call Create_Events.CreateEvents
    Application.StatusBar = "Liquidating UTXOs..." & year
    Call Liquidate_Events.Liquidate
End Sub
