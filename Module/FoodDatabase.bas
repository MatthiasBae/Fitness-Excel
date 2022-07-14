Attribute VB_Name = "FoodDatabase"
Public Function GetFoods(Optional Name As String, Optional Brand As String, Optional TopCount As Integer, Optional RecipesOnly As Boolean = False) As Dictionary
    Dim FoodIdRange As Range, rng As Range, FoodList As New Dictionary
    Dim ItemCount As Integer
    Dim SelectedFood As Food
    
    Dim Tbl As ListObject
    Set Tbl = FoodConfigs.FoodTable
    
    On Error GoTo Err:
    
    If Name <> "" Then
        Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Nahrungsmittel").Index, Criteria1:="=*" & Name & "*", _
            Operator:=xlAnd
    End If
    
    If Brand <> "" Then
        Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Marke").Index, Criteria1:="=*" & Brand & "*", _
            Operator:=xlAnd
    End If
    
    If RecipesOnly = True Then
        Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Rezept").Index, Criteria1:="=" & True, _
            Operator:=xlAnd
    End If
    
    Set FoodIdRange = Tbl.ListColumns("NahrungsmittelId").DataBodyRange.SpecialCells(xlCellTypeVisible)
    ItemCount = 1
    For Each rng In FoodIdRange
        
        Set SelectedFood = New Food
        SelectedFood.Load rng.Value
        
        FoodList.Add SelectedFood.FoodId, SelectedFood
        If ItemCount >= TopCount Then
            Exit For
        End If
        ItemCount = ItemCount + 1
    Next rng
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Nahrungsmittel").Index
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Marke").Index
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Rezept").Index
    
    Set GetFoods = FoodList
    Exit Function
Err:
    Set GetFoods = FoodList
End Function

Public Function FoodExists(Name As String, Brand As String)
    Dim Tbl As ListObject
    Dim FoodCount As Integer
    
    Set Tbl = FoodConfigs.FoodTable
    
    FoodCount = WorksheetFunction.CountIfs(Tbl.ListColumns("Nahrungsmittel").Range, Name, Tbl.ListColumns("Marke").Range, Brand)
    FoodExists = IIf(FoodCount > 0, True, False)
End Function
Public Function FoodUnitExists(FoodId As Long, UnitName As String) As Range
    Dim Tbl As ListObject
    Dim UnitId As String
    
    Dim IdRange As Range
    
    Set Tbl = FoodConfigs.FoodUnitsTable
    
    UnitId = Replace(Str(FoodId) & UnitName, " ", "")
    Set IdRange = Tbl.ListColumns("Id").Range.Find(what:=UnitId, lookat:=xlWhole, LookIn:=xlValues)
    'FoodCount = WorksheetFunction.CountIfs(Tbl.ListColumns("NahrungsmittelId").Range, FoodId, Tbl.ListColumns("Einheit").Range, UnitName)
    Set FoodUnitExists = IdRange
End Function

Public Function GetFoodId(Name As String, Brand As String)
    Dim Tbl As ListObject
    Set Tbl = FoodConfigs.FoodTable
    
    GetFoodId = WorksheetFunction.SumIfs(Tbl.ListColumns("NahrungsmittelId").Range, Tbl.ListColumns("Nahrungsmittel").Range, Name, Tbl.ListColumns("Marke").Range, Brand)
End Function

Public Function GetMaxFoodId() As Long
    Dim Tbl As ListObject
    Set Tbl = FoodConfigs.FoodTable
    GetMaxFoodId = WorksheetFunction.Max(Tbl.ListColumns("NahrungsmittelId").Range)
End Function

Public Sub SaveFood(Name As String, Brand As String, UnitName As String, Amount As Double, Calories As Double, Protein As Double, Carbs As Double, Sugar As Double, Fat As Double, Optional ACG1 As String, Optional ACG2 As String, Optional ACG3 As String)
    Dim FoodTbl As ListObject, FoodUnitTbl As ListObject
    
    Dim FoodRow As Range
    Dim FoodDataset As ListRow
    Dim FoodId As Long
    
    FoodId = FoodDatabase.GetFoodId(Name, Brand)
    
    Set FoodTbl = FoodConfigs.FoodTable
    If FoodId > 0 Then
        Set FoodRow = FoodTbl.ListColumns("NahrungsmittelId").DataBodyRange.Find(what:=FoodId, LookIn:=xlValues, lookat:=xlWhole)
        Set FoodDataset = FoodTbl.ListRows(FoodRow.Row - 1)
    Else
        Set FoodDataset = FoodTbl.ListRows.Add
        FoodId = GetMaxFoodId + 1
    End If

    FoodTbl.ListColumns("NahrungsmittelId").DataBodyRange(FoodDataset.Index) = FoodId
    FoodTbl.ListColumns("Nahrungsmittel").DataBodyRange(FoodDataset.Index) = Name
    FoodTbl.ListColumns("Marke").DataBodyRange(FoodDataset.Index) = Brand
    FoodTbl.ListColumns("KategorieLevel1").DataBodyRange(FoodDataset.Index) = ACG1
    FoodTbl.ListColumns("KategorieLevel2").DataBodyRange(FoodDataset.Index) = ACG2
    FoodTbl.ListColumns("KategorieLevel3").DataBodyRange(FoodDataset.Index) = ACG3
    
    FoodDatabase.SaveFoodUnit FoodId, UnitName, Amount, Calories, Protein, Carbs, Sugar, Fat
    
End Sub

Public Sub SaveFoodUnit(FoodId As Long, UnitName As String, Amount As Double, Calories As Double, Protein As Double, Carbs As Double, Sugar As Double, Fat As Double, Optional IsDefault As Boolean = False)
    Dim FoodUnitRange As Range
    Dim FoodUnitDataset As ListRow
    
    Set FoodUnitTbl = FoodConfigs.FoodUnitsTable

    Set FoodUnitRange = FoodDatabase.FoodUnitExists(FoodId, UnitName)
    If FoodUnitRange Is Nothing Then
        Set FoodUnitDataset = FoodUnitTbl.ListRows.Add
    Else
        Set FoodUnitDataset = FoodUnitTbl.ListRows(FoodUnitRange.Row - 1)
    End If
    
    FoodUnitTbl.ListColumns("NahrungsmittelId").DataBodyRange(FoodUnitDataset.Index) = FoodId
    FoodUnitTbl.ListColumns("Einheit").DataBodyRange(FoodUnitDataset.Index) = UnitName
    FoodUnitTbl.ListColumns("Standardeinheit").DataBodyRange(FoodUnitDataset.Index) = IsDefault
    FoodUnitTbl.ListColumns("Menge").DataBodyRange(FoodUnitDataset.Index) = Amount
    FoodUnitTbl.ListColumns("Kalorien").DataBodyRange(FoodUnitDataset.Index) = Calories
    FoodUnitTbl.ListColumns("Proteine").DataBodyRange(FoodUnitDataset.Index) = Protein
    FoodUnitTbl.ListColumns("Kohlenhydrate").DataBodyRange(FoodUnitDataset.Index) = Carbs
    FoodUnitTbl.ListColumns("Zucker").DataBodyRange(FoodUnitDataset.Index) = Sugar
    FoodUnitTbl.ListColumns("Fett").DataBodyRange(FoodUnitDataset.Index) = Fat
    
End Sub
