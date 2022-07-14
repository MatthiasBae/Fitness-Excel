VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SummaryDate As Date
Public CaloriesEaten As Double
Public CaloriesBurned As Double
Public CaloriesDifference As Double
Public Weight As Double
Public Fat As Double
Public Steps As Long
Public Distance As Double
Public Heartrate As Integer
Public Worked As Boolean
Public Homeoffice As Boolean

Public Sub Load()
    Dim Tbl As ListObject
    Dim Result As Range

    Set Tbl = SummaryConfigs.SummaryTable
    Set Result = Tbl.ListColumns("Datum").Range.Find(what:=SelectedDate, LookIn:=xlValues, lookat:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No Summary found with for date {0}", SelectedDate)
        Exit Sub
    End If
    
    Me.SummaryDate = Tbl.ListColumns("Datum").Range.Rows(Result.Row).Value
    Me.CaloriesEaten = Tbl.ListColumns("Kalorien gegessen").Range.Rows(Result.Row).Value
    Me.CaloriesBurned = Tbl.ListColumns("Kalorien Aktivitätsbedarf").Range.Rows(Result.Row).Value
    Me.CaloriesDifference = Tbl.ListColumns("Kalorien Differenz").Range.Rows(Result.Row).Value
    Me.Weight = Round(Tbl.ListColumns("Gewicht").Range.Rows(Result.Row).Value, 1)
    Me.Fat = Round(Tbl.ListColumns("Fett").Range.Rows(Result.Row).Value, 1)
    Me.Steps = Round(Tbl.ListColumns("Schritte").Range.Rows(Result.Row).Value, 1)
    Me.Distance = Round(Tbl.ListColumns("Distanz").Range.Rows(Result.Row).Value, 1)
    Me.Heartrate = Round(Tbl.ListColumns("Herzfrequenz").Range.Rows(Result.Row).Value, 1)
    Me.Worked = Round(Tbl.ListColumns("Gearbeitet").Range.Rows(Result.Row).Value, 1)
    Me.Homeoffice = Round(Tbl.ListColumns("Homeoffice").Range.Rows(Result.Row).Value, 1)
End Sub
