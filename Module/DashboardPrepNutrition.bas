Attribute VB_Name = "DashboardPrepNutrition"

Public Sub LoadNutrientsDivision()
    Dim Ws As Worksheet, WsRawdata As Worksheet
    Dim Tbl As ListObject
    Dim DateSearchField As Range
    
    Dim Result As Range
    
    On Error GoTo Err:
    
    Set Ws = Worksheets("Vorbereitung Ernährungsplan")
    Set DateSearchField = Ws.Range("TextDateSearchField")
    
    Set WsRawdata = Worksheets("Rohdaten_Nährstoffverteilung")
    Set Tbl = WsRawdata.ListObjects("TblNutrientDivision")
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum von").Range.Column, Criteria1:="<=" & CLng(DateSearchField.Value) & "", Operator:=xlAnd
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum bis").Range.Column, Criteria1:=">=" & CLng(DateSearchField.Value) & "", Operator:=xlAnd
    
    Set Result = Tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
    If Result.Rows.Count > 1 Then
        Debug.Print "Zeiträume in der Nährstoffverteilung falsch gepflegt"
        Exit Sub
    End If
    
    Ws.Range("TextNutrientDivisionDateFrom").Value = Tbl.ListColumns("Datum von").Range.Rows(Result.Row).Value
    Ws.Range("TextNutrientDivisionDateTo").Value = Tbl.ListColumns("Datum bis").Range.Rows(Result.Row).Value
    Ws.Range("TextNutrientDivisionCalories").Value = Tbl.ListColumns("Kalorien in Kcal.").Range.Rows(Result.Row).Value
    Ws.Range("TextNutrientDivisionProtein").Value = Tbl.ListColumns("Proteine in %").Range.Rows(Result.Row).Value
    Ws.Range("TextNutrientDivisionCarbs").Value = Tbl.ListColumns("Kohlenhydrate in %").Range.Rows(Result.Row).Value
    Ws.Range("TextNutrientDivisionFat").Value = Tbl.ListColumns("Fett in %").Range.Rows(Result.Row).Value
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum von").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum bis").Range.Column
    Exit Sub
Err:
    reset
End Sub

Public Sub SaveNutrientsDivision()

End Sub

Public Sub reset()
    Dim Ws As Worksheet
    Set Ws = Worksheets("Vorbereitung Ernährungsplan")
    
    Ws.Range("TextNutrientDivisionDateFrom").Value = ""
    Ws.Range("TextNutrientDivisionDateTo").Value = ""
    Ws.Range("TextNutrientDivisionCalories").Value = ""
    Ws.Range("TextNutrientDivisionProtein").Value = ""
    Ws.Range("TextNutrientDivisionCarbs").Value = ""
    Ws.Range("TextNutrientDivisionFat").Value = ""
End Sub
