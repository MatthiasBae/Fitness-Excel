Attribute VB_Name = "ButtonEvents"

Public Sub FoodButton_Click(FoodId As Long)
    Dim FoodItem As New Food
    FoodItem.Load FoodId
    
    If ActiveSheet.Name = "Dashboard Lebensmittel" Then
        Set FoodDashboard.SelectedFood = FoodItem
        FoodDashboard.FillSelectedFoodPanel FoodDashboard.SelectedFood
        
    ElseIf ActiveSheet.Name = "Dashboard Ernährung" Then
        Set NutritionDashboard.SelectedFood = FoodItem
        NutritionDashboard.FillSelectedFoodPanel NutritionDashboard.SelectedFood
        
    End If
End Sub

Public Sub PlanMealButton_Click(MealId As Integer)
    Dim MealItem As New NutritionPlanMeal
    
    If NutritionDashboard.SelectedPlan Is Nothing Then
        Set NutritionDashboard.SelectedPlan = New NutritionPlan
        NutritionDashboard.SelectedPlan.Load Worksheets("Dashboard Ernährung").Range("Text_Nt_DateFrom")
    End If
    
    Set MealItem = NutritionDashboard.SelectedPlan.Meals(MealId)
    
    Set NutritionDashboard.SelectedPlanMeal = MealItem
    NutritionDashboard.FillPlanMealFoodList NutritionDashboard.SelectedPlanMeal
End Sub

Public Sub PlanMealButton_Delete_Click(MealId As Integer)
    Dim MealItem As New NutritionPlanMeal
    
    If NutritionDashboard.SelectedPlan Is Nothing Then
        Set NutritionDashboard.SelectedPlan = New NutritionPlan
        NutritionDashboard.SelectedPlan.Load Worksheets("Dashboard Ernährung").Range("Text_Nt_DateFrom")
    End If
    
    Set MealItem = NutritionDashboard.SelectedPlan.Meals(MealId)
    MealItem.Delete
    NutritionDashboard.SelectedPlan.Meals.Remove MealItem.Id
    NutritionDashboard.FillPlanMealList NutritionDashboard.SelectedPlan
End Sub
