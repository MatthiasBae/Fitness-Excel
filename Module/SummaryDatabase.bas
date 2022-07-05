Attribute VB_Name = "SummaryDatabase"
Public Sub Update()
    Dim Tbl As ListObject
    Dim LastDate As Date, Today As Date, CurrentDate As Date
    Dim Dataset As ListRow
    Dim i As Integer
    
    Set Tbl = SummaryConfigs.SummaryTable
    LastDate = WorksheetFunction.Max(Tbl.ListColumns("Datum").Range)
    Today = Date
    
    For i = 1 To Functions.DateDiff(LastDate, Today)
        CurrentDate = DateAdd(LastDate, i)
        
        Set Dataset = Tbl.ListRows.Add
        Tbl.ListColumns("Datum").DataBodyRange(Dataset.Index) = CurrentDate
    Next i
    Tbl.ListColumns("Datum").DataBodyRange.NumberFormatLocal = "tt.MM.jjjj"
End Sub
