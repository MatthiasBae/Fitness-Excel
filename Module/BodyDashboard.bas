Attribute VB_Name = "BodyDashboard"

Public Sub FillBodyList()
    Dim Ws As Worksheet
    Dim DateFrom As Date, WeightFilter As String, FatFilter As String
    
    Set Ws = Worksheets("Dashboard Körper")
    
    On Error GoTo Err:
    
    DateFrom = IIf(Ws.Range("TextSearchDateFromField") = "", Date, Ws.Range("TextSearchDateFromField"))
    WeightFilter = Ws.Range("TextSearchWeightField")
    FatFilter = Ws.Range("TextSearchFatField")
    
    ResetBodyList
    BodyDatabase.FillBodyList Ws.Range("ListBodies"), DateFrom, WeightFilter, FatFilter
    Application.CutCopyMode = False
    Exit Sub
Err:
    ResetBodyList
End Sub

Public Sub ResetBodyList()
    Dim Shp As Shape
    For Each Shp In Worksheets("Dashboard Körper").Shapes
        If InStr(1, Shp.Name, "BtnBody_") Then
            Shp.Delete
        End If
    Next Shp
End Sub
Public Sub DeleteBody()
    Dim Ws As Worksheet
    Set Ws = Worksheets("Dashboard Körper")
    
    Dim SelectedButtonName As String, SelectedButtonId As String
    Dim SelectedBody As New Body
    Dim SelectedBodyDate As Date
    
    SelectedButtonName = Application.Caller
    SelectedButtonId = Functions.GetIdFromButton(SelectedButtonName)
    SelectedBodyDate = CDate(Right(SelectedButtonId, 2) & "." & Mid(SelectedButtonId, 5, 2) & "." & Left(SelectedButtonId, 4))
    
    SelectedBody.Load SelectedBodyDate
    
    SelectedBody.Delete
    FillBodyList
    
End Sub
