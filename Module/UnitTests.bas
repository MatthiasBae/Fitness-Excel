Attribute VB_Name = "UnitTests"
Public Function IsNutritionPlanLoading() As Boolean
    Dim SelectedDate As Date
    Dim Plan As New NutritionPlan
    
    On Error GoTo Err:
    
    SelectedDate = Date
    
    Plan.Load SelectedDate
    IsNutritionPlanLoading = True
    Exit Function
Err:
    IsNutritionPlanLoading = False
    Debug.Print Err.Description
End Function

Public Function CanAddFoodToNutritionPlan() As Boolean
    Dim DateFrom As Date, DateTo As Date
    Dim Food As New Food
    
    On Error GoTo Err:
    Food.Load 2
    
    CanAddFoodToNutritionPlan = NutritionPlanDatabase.TryAddFood(Food, "Gramm", 150, 2, Date, Date + 6, True, 2)
    Exit Function
Err:
    CanAddFoodToNutritionPlan = False
    Debug.Print (Err.Description)
End Function

Public Function CanLoadBody() As Boolean
    Dim Bd As New Body
    On Error GoTo Err:
    Bd.Load Date
    
    CanLoadBody = True
    Exit Function
Err:
    CanLoadBody = False
    Debug.Print Err.Description
End Function
Public Function CanLoadFood() As Boolean
    Dim fd As New Food
    On Error GoTo Err:
    fd.Load 18
    
    CanLoadFood = True
    Exit Function
Err:
    CanLoadFood = False
    Debug.Print Err.Description
End Function
