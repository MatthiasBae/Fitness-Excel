VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NutritionPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public PlanDate As Date
Public Meals As New Dictionary

Public Sub Load(SelectedDate As Date)
    Me.Meals.RemoveAll
    Me.LoadPlan SelectedDate
    Me.LoadPlanMeals SelectedDate
End Sub

Public Sub LoadPlan(SelectedDate As Date)
    Dim Tbl As ListObject
    Dim Result As Range

    Set Tbl = NutritionConfigs.MealTable
    Set Result = Tbl.ListColumns("Datum").Range.Find(what:=SelectedDate, LookIn:=xlValues, lookat:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No Plan found with for date {0}", SelectedDate)
        Exit Sub
    End If
    
    Me.PlanDate = Tbl.ListColumns("Datum").Range.Rows(Result.Row).Value

End Sub
Public Sub LoadPlanMeals(SelectedDate As Date)
    Dim Tbl As ListObject
    Dim Result As Range, SearchRange As Range, FirstFoundAddress As String
    
    Dim Meal As NutritionPlanMeal

    Set Tbl = NutritionConfigs.MealTable
    Set SearchRange = Tbl.ListColumns("Datum").Range
    Set Result = SearchRange.Find(what:=SelectedDate, LookIn:=xlValues, lookat:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No PlanMeal found for Date {0}", SelectedDate)
        Exit Sub
    End If
    
    FirstFoundAddress = Result.Address
    
    Do
        Set Meal = New NutritionPlanMeal
        
        Meal.Id = Tbl.ListColumns("MahlzeitenId").Range(Result.Row).Value
        Meal.MealType = Tbl.ListColumns("MahlzeitenTyp").Range(Result.Row).Value
        
        Meal.LoadMealFoods (SelectedDate)
        Me.TryAddMeal Meal.Id, Meal
        
        Set Result = SearchRange.FindNext(Result)
    Loop While Not Result Is Nothing And FirstFoundAddress <> Result.Address
    
    Set Me.Meals = Functions.SortDictionaryByKey(Me.Meals, xlAscending)
End Sub
Public Function TryAddMeal(MealId As Integer, Meal As NutritionPlanMeal) As Boolean

    If Me.Meals.Exists(MealId) Then
        TryAddMeal = False
        Exit Function
    End If
    
    Me.AddMeal MealId, Meal
    TryAddMeal = True
End Function

Public Sub AddMeal(MealId As Integer, Meal As NutritionPlanMeal)
    Me.Meals.Add MealId, Meal
End Sub


Public Function GetNutrients() As Nutrients
    Dim NewNutrients As New Nutrients
    Dim MealNutrients As New Nutrients
    Dim Key As Variant
    For Each Key In Me.Meals.Keys
        Set MealNutrients = Meals(Key).GetNutrients
        NewNutrients.Calories = NewNutrients.Calories + MealNutrients.Calories
        NewNutrients.Macros.Protein = NewNutrients.Macros.Protein + MealNutrients.Macros.Protein
        NewNutrients.Macros.Carbohydrates = NewNutrients.Macros.Carbohydrates + MealNutrients.Macros.Carbohydrates
        NewNutrients.Macros.Sugar = NewNutrients.Macros.Sugar + MealNutrients.Macros.Sugar
        NewNutrients.Macros.Fat = NewNutrients.Macros.Fat + MealNutrients.Macros.Fat
        
        NewNutrients.Micros.Fiber = NewNutrients.Micros.Fiber + MealNutrients.Micros.Fiber
    Next Key
    Set GetNutrients = NewNutrients
End Function

Public Function GetButton() As Shape
    Dim Shp As Shape, ShpDup As Shape
    Dim DayNutrients As New Nutrients
    
    Set DayNutrients = Me.GetNutrients
    
    Set Shp = Worksheets("Templates").Shapes("BtnTemplate_Plan")
    
    Set ShpDup = Shp.Duplicate

    ShpDup.GroupItems("BtnTemplate_Plan_Date").TextEffect.Text = Printf("{0}", Me.PlanDate)
    ShpDup.GroupItems("BtnTemplate_Plan_MealCount").TextEffect.Text = Printf("Mahlzeiten: {0} Stück", Me.Meals.Count)
    ShpDup.GroupItems("BtnTemplate_Plan_Macros").TextEffect.Text = Printf("{0} Kcal. - (P: {1} | K: {2} | F: {3})", DayNutrients.Calories, DayNutrients.Macros.Protein, DayNutrients.Macros.Carbohydrates, DayNutrients.Macros.Fat)
    
    ShpDup.GroupItems("BtnTemplate_Plan_Date").Name = Printf("BtnPlan_Date_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    ShpDup.GroupItems("BtnTemplate_Plan_MealCount").Name = Printf("BtnPlan_MealCount_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    ShpDup.GroupItems("BtnTemplate_Plan_Macros").Name = Printf("BtnPlan_Macros_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    ShpDup.Name = Printf("BtnPlan_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    
    Set GetButton = ShpDup
End Function
