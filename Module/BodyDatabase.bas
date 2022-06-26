Attribute VB_Name = "BodyDatabase"
Public Sub FillBodyList(SelectedRng As Range, DateFrom As Date, Optional WeightFilter As String = "", Optional FatFilter As String = "")
    Dim WrapPanel As New WrapPanel
    Dim BodyList As New Dictionary, BodyId As Variant, SelectedBody As Body
    Dim BodyBtn As Shape
    
    Set BodyList = BodyDatabase.GetBodies(DateFrom, WeightFilter, FatFilter)
    
    WrapPanel.Initialize SelectedRng, 1
    
    Application.ScreenUpdating = False
    For Each BodyId In BodyList.Keys
        Set SelectedBody = BodyList(BodyId)
        Set BodyBtn = SelectedBody.GetButton
        
        WrapPanel.Add BodyBtn
    Next
    WrapPanel.Render
    Application.ScreenUpdating = True
End Sub

Public Function GetBodies(Optional DateFrom As Date, Optional WeightFilter As String = "", Optional FatFilter As String = "") As Dictionary
    Dim Ws As Worksheet
    Dim BodyIdRange As Range, Rng As Range, BodyList As New Dictionary
    Dim ItemCount As Integer
    Dim SelectedBody As Body
    
    Dim Tbl As ListObject
    
    Set Ws = Worksheets(Configs.BodyWorksheetName)
    Set Tbl = Ws.ListObjects("TblBody")
    
    Tbl.Sort.SortFields.Clear
    Tbl.Sort.SortFields.Add2 Key:=Tbl.ListColumns("Datum").Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With Tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    On Error GoTo Err:
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column, Criteria1:=">" & CLng(DateFrom), _
            Operator:=xlAnd
    
    If WeightFilter <> "" Then
        Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Gewicht").Range.Column, Criteria1:=WeightFilter, _
            Operator:=xlAnd
    End If
    If FatFilter <> "" Then
        Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Fett").Range.Column, Criteria1:=FatFilter, _
            Operator:=xlAnd
    End If
    
    
    
    Set BodyIdRange = Tbl.ListColumns("Datum").DataBodyRange.SpecialCells(xlCellTypeVisible)
    ItemCount = 1
    For Each Rng In BodyIdRange
        
        Set SelectedBody = New Body
        SelectedBody.Load Rng.Value
        
        BodyList.Add SelectedBody.PlanDate, SelectedBody
    Next Rng
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Gewicht").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Fett").Range.Column
    Set GetBodies = BodyList
    Exit Function
Err:
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Gewicht").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Fett").Range.Column
End Function

