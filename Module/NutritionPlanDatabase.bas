Attribute VB_Name = "NutritionPlanDatabase"
'@TODO: Refactoring
Public Function GetPlans(DateFrom As Date, DateTo As Date) As Dictionary
    Dim PlanDateRange As Range, rng As Range, PlanList As New Dictionary
    Dim SelectedPlan As NutritionPlan
    
    Dim Tbl As ListObject
    
    Set Tbl = NutritionConfigs.MealTable
    
    Tbl.Sort.SortFields.Clear
    Tbl.Sort.SortFields.Add2 Key:=Tbl.ListColumns("Datum").Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With Tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Tbl.Sort.SortFields.Clear
    Tbl.Sort.SortFields.Add2 Key:=Tbl.ListColumns("Datum").Range, SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    Tbl.Sort.SortFields.Add2 Key:=Tbl.ListColumns("MahlzeitenId").Range, SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With Tbl _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '@TODO: Filtert alles aus wenn ich über VBA den Filter setze.
    'Bestätige ich den Filter nochmal mit "OK" im Tabellenblatt, funktioniert der
    'Tbl.Range.AutoFilter Field:=1, Criteria1:=">=" & CLng(DateFrom) & "", Operator:=xlAnd, Criteria2:="<=" & CLng(DateTo) & ""
    Tbl.Range.AutoFilter Field:=1, Criteria1:="=" & DateFrom & ""

    Set PlanDateRange = Tbl.ListColumns("Datum").DataBodyRange.SpecialCells(xlCellTypeVisible)
    For Each rng In PlanDateRange

        If Not PlanList.Exists(rng.Value) Then
            Set SelectedPlan = New NutritionPlan
            SelectedPlan.Load rng.Value
            
            PlanList.Add SelectedPlan.PlanDate, SelectedPlan
        End If
    Next rng
    
    Tbl.Range.AutoFilter Field:=1
    Set GetPlans = PlanList
End Function

Public Function PlanMealExists(SelectedDate As Date, MealId As Integer) As Boolean
    Dim i As Long
    Dim Tbl As ListObject
    
    Set Tbl = NutritionConfigs.MealTable
    i = WorksheetFunction.CountIfs(Tbl.ListColumns("Datum").Range, SelectedDate, Tbl.ListColumns("MahlzeitenId").Range, MealId)
    PlanMealExists = IIf(i > 0, True, False)
End Function
Public Function PlanMealFoodExists(SelectedDate As Date, MealId As Integer, FoodId As Long) As Boolean
    Dim i As Long
    Dim Tbl As ListObject
    
    Set Tbl = NutritionConfigs.MealFoodTable
    i = WorksheetFunction.CountIfs(Tbl.ListColumns("Datum").Range, SelectedDate, Tbl.ListColumns("MahlzeitenId").Range, MealId, Tbl.ListColumns("NahrungsmittelId").Range, FoodId)
    PlanMealFoodExists = IIf(i > 0, True, False)
End Function

Public Function TryAddFood(FoodItem As Food, UnitName As String, Amount As Double, MealId As Integer, DateFrom As Date, DateTo As Date, Optional IsCheatMeal As Boolean = False, Optional ExceptWeekday As Integer = 0) As Boolean
    Dim MealTbl As ListObject, MealFoodTbl As ListObject
    Dim MealDataset As ListRow, MealFoodDataset As ListRow
    
    Dim MealExists As Boolean, MealFoodExists As Boolean
    
    Dim i As Integer
    Dim CurrentDate As Date
    
    Set MealTbl = NutritionConfigs.MealTable
    Set MealFoodTbl = NutritionConfigs.MealFoodTable
    
    For i = 0 To DateDiff(DateFrom, DateTo)
        CurrentDate = DateAdd(DateFrom, i)
                
        If Weekday(CurrentDate, vbMonday) = ExceptWeekday Then Exit For
        
        MealExists = NutritionPlanDatabase.PlanMealExists(CurrentDate, MealId)
        MealFoodExists = NutritionPlanDatabase.PlanMealFoodExists(CurrentDate, MealId, FoodItem.FoodId)

        If Not MealExists Then
            Set MealDataset = MealTbl.ListRows.Add
            
            MealTbl.ListColumns("Datum").DataBodyRange(MealDataset.Index) = CurrentDate
            MealTbl.ListColumns("MahlzeitenId").DataBodyRange(MealDataset.Index) = MealId
            'MealTbl.ListColumns("Datum").Range(MealDataset.index) = CurrentDate
            MealTbl.ListColumns("Cheatmeal").DataBodyRange(MealDataset.Index) = IsCheatMeal
        End If
        
        If Not MealFoodExists Then
            Set MealFoodDataset = MealFoodTbl.ListRows.Add
            
            MealFoodTbl.ListColumns("Datum").DataBodyRange(MealFoodDataset.Index) = CurrentDate
            MealFoodTbl.ListColumns("MahlzeitenId").DataBodyRange(MealFoodDataset.Index) = MealId
            MealFoodTbl.ListColumns("NahrungsmittelId").DataBodyRange(MealFoodDataset.Index) = FoodItem.FoodId
            MealFoodTbl.ListColumns("Geplant").DataBodyRange(MealFoodDataset.Index) = IIf(CurrentDate > Date, True, False)
            MealFoodTbl.ListColumns("Nahrungsmittel").DataBodyRange(MealFoodDataset.Index) = FoodItem.Name
            MealFoodTbl.ListColumns("Marke").DataBodyRange(MealFoodDataset.Index) = FoodItem.Brand
            MealFoodTbl.ListColumns("Einheit").DataBodyRange(MealFoodDataset.Index) = UnitName
            MealFoodTbl.ListColumns("Menge").DataBodyRange(MealFoodDataset.Index) = Amount
        End If
    Next i

    MealTbl.ListColumns("Datum").DataBodyRange.NumberFormatLocal = "tt.MM.jjjj"
    MealFoodTbl.ListColumns("Datum").DataBodyRange.NumberFormatLocal = "tt.MM.jjjj"
End Function

Public Sub DeleteMeal(PlanMeal As NutritionPlanMeal)
    Dim FoodKey As Variant, FoodItem As New NutritionPlanMealFood

    Dim Tbl As ListObject
    
    Dim rng As Range
    
    For Each FoodKey In PlanMeal.Foods.Keys
        Set FoodItem = PlanMeal.Foods(FoodKey)
        FoodItem.Delete
    Next FoodKey
    
    PlanMeal.Foods.RemoveAll

    Set Tbl = NutritionConfigs.MealTable
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column, Criteria1:="=" & PlanMeal.MealDate & ""
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column, Criteria1:="=" & PlanMeal.Id & ""
    
    Application.DisplayAlerts = False
    For Each rng In Tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
        rng.EntireRow.Delete
    Next rng
    Application.DisplayAlerts = True
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column
    
End Sub

Public Sub DeleteMealFood(PlanMealFood As NutritionPlanMealFood)
    Dim Tbl As ListObject
    Dim rng As Range

    Set Tbl = NutritionConfigs.MealFoodTable
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column, Criteria1:="=" & PlanMealFood.PlanDate & ""
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column, Criteria1:="=" & PlanMealFood.MealId & ""
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("NahrungsmittelId").Range.Column, Criteria1:="=" & PlanMealFood.FoodId & ""
    
    Application.DisplayAlerts = False
    For Each rng In Tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
        rng.EntireRow.Delete
    Next rng
    Application.DisplayAlerts = True
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("NahrungsmittelId").Range.Column
End Sub
