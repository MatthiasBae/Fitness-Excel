Attribute VB_Name = "FoodDatabase"
Public Function GetFoods(Name As String, Brand As String, Optional TopCount As Integer) As Dictionary
    Dim FoodIdRange As Range, Rng As Range, FoodList As New Dictionary
    Dim ItemCount As Integer
    Dim SelectedFood As Food
    
    Dim Tbl As ListObject
    Set Tbl = FoodConfigs.FoodTable
        
    Tbl.Range.AutoFilter Field:=11, Criteria1:="=*" & Name & "*", _
        Operator:=xlAnd
    Tbl.Range.AutoFilter Field:=12, Criteria1:="=*" & Brand & "*", _
        Operator:=xlAnd

    Set FoodIdRange = Tbl.ListColumns("NahrungsmittelId").DataBodyRange.SpecialCells(xlCellTypeVisible)
    ItemCount = 1
    For Each Rng In FoodIdRange
        
        Set SelectedFood = New Food
        SelectedFood.Load Rng.Value
        
        FoodList.Add SelectedFood.FoodId, SelectedFood
        If ItemCount >= TopCount Then
            Exit For
        End If
        ItemCount = ItemCount + 1
    Next Rng
    
    Tbl.Range.AutoFilter Field:=11
    Tbl.Range.AutoFilter Field:=12
    
    Set GetFoods = FoodList
End Function

