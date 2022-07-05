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
    Dim BodyIdRange As Range, rng As Range, BodyList As New Dictionary
    Dim SelectedBody As Body
    Dim Tbl As ListObject

    Set Tbl = BodyConfigs.BodyTable
    
    If BodyExists(DateFrom, WeightFilter, FatFilter) = False Then
        Set GetBodies = New Dictionary
        Exit Function
    End If
    
    Tbl.Sort.SortFields.Clear
    Tbl.Sort.SortFields.Add2 Key:=Tbl.ListColumns("Datum").Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With Tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

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

    For Each rng In BodyIdRange
        Set SelectedBody = New Body
        SelectedBody.Load rng.Value
        
        BodyList.Add SelectedBody.PlanDate, SelectedBody
    Next rng
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Gewicht").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Fett").Range.Column
    Set GetBodies = BodyList

End Function

Public Function BodyExists(Optional DateFrom As Date, Optional WeightFilter As String = "", Optional FatFilter As String = "")
    Dim i As Long
    Dim Tbl As ListObject
    Set Tbl = BodyConfigs.BodyTable
    
    i = WorksheetFunction.CountIfs(Tbl.ListColumns("Datum").Range, ">" & CLng(DateFrom), Tbl.ListColumns("Gewicht").Range, WeightFilter, Tbl.ListColumns("Fett").Range, FatFilter)
    BodyExists = IIf(i > 0, True, False)
    
End Function
