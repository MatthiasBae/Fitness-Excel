Attribute VB_Name = "NutritionPlanDatabase"
'@TODO: Refactoring
Public Function GetPlans(DateFrom As Date, DateTo As Date) As Dictionary
    Dim PlanDateRange As Range, Rng As Range, PlanList As New Dictionary
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
    For Each Rng In PlanDateRange

        If Not PlanList.Exists(Rng.Value) Then
            Set SelectedPlan = New NutritionPlan
            SelectedPlan.Load Rng.Value
            
            PlanList.Add SelectedPlan.PlanDate, SelectedPlan
        End If
    Next Rng
    
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
            
            MealTbl.ListColumns("Datum").DataBodyRange(MealDataset.index) = CurrentDate
            MealTbl.ListColumns("MahlzeitenId").DataBodyRange(MealDataset.index) = MealId
            'MealTbl.ListColumns("Datum").Range(MealDataset.index) = CurrentDate
            MealTbl.ListColumns("Cheatmeal").DataBodyRange(MealDataset.index) = IsCheatMeal
        End If
        
        If Not MealFoodExists Then
            Set MealFoodDataset = MealFoodTbl.ListRows.Add
            
            MealFoodTbl.ListColumns("Datum").DataBodyRange(MealFoodDataset.index) = CurrentDate
            MealFoodTbl.ListColumns("MahlzeitenId").DataBodyRange(MealFoodDataset.index) = MealId
            MealFoodTbl.ListColumns("NahrungsmittelId").DataBodyRange(MealFoodDataset.index) = FoodItem.FoodId
            MealFoodTbl.ListColumns("Geplant").DataBodyRange(MealFoodDataset.index) = IIf(CurrentDate > Date, True, False)
            MealFoodTbl.ListColumns("Nahrungsmittel").DataBodyRange(MealFoodDataset.index) = FoodItem.Name
            MealFoodTbl.ListColumns("Marke").DataBodyRange(MealFoodDataset.index) = FoodItem.Brand
            MealFoodTbl.ListColumns("Einheit").DataBodyRange(MealFoodDataset.index) = UnitName
            MealFoodTbl.ListColumns("Menge").DataBodyRange(MealFoodDataset.index) = Amount
        End If
    Next i

    MealTbl.ListColumns("Datum").DataBodyRange.NumberFormatLocal = "tt.MM.jjjj"
    MealFoodTbl.ListColumns("Datum").DataBodyRange.NumberFormatLocal = "tt.MM.jjjj"
End Function

Public Sub DeleteMeal(PlanMeal As NutritionPlanMeal)
    Dim FoodKey As Variant, FoodItem As New NutritionPlanMealFood

    Dim Tbl As ListObject
    
    Dim Rng As Range
    
    For Each FoodKey In PlanMeal.Foods.Keys
        Set FoodItem = PlanMeal.Foods(FoodKey)
        FoodItem.Delete
    Next FoodKey
    
    PlanMeal.Foods.RemoveAll

    Set Tbl = NutritionConfigs.MealTable
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column, Criteria1:="=" & PlanMeal.MealDate & ""
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column, Criteria1:="=" & PlanMeal.Id & ""
    
    Application.DisplayAlerts = False
    For Each Rng In Tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
        Rng.EntireRow.Delete
    Next Rng
    Application.DisplayAlerts = True
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column
    
End Sub

Public Sub DeleteMealFood(PlanMealFood As NutritionPlanMealFood)
    Dim Tbl As ListObject
    Dim Rng As Range

    Set Tbl = NutritionConfigs.MealFoodTable
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column, Criteria1:="=" & PlanMealFood.PlanDate & ""
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column, Criteria1:="=" & PlanMealFood.MealId & ""
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("NahrungsmittelId").Range.Column, Criteria1:="=" & PlanMealFood.FoodId & ""
    
    Application.DisplayAlerts = False
    For Each Rng In Tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
        Rng.EntireRow.Delete
    Next Rng
    Application.DisplayAlerts = True
    
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("Datum").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("MahlzeitenId").Range.Column
    Tbl.Range.AutoFilter Field:=Tbl.ListColumns("NahrungsmittelId").Range.Column
End Sub
