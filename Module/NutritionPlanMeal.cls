VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NutritionPlanMeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MealDate As Date
Public Id As Integer

Public MealType As String
Public Foods As New Dictionary

Public Sub LoadMealFoods(SelectedDate As Date)
    Dim Tbl As ListObject
    Dim Result As Range, SearchRange As Range, FirstFoundAddress As String
    
    Dim Food As NutritionPlanMealFood
    Dim MealId As Integer
    
    Me.MealDate = SelectedDate

    Set Tbl = NutritionConfigs.MealFoodTable
    Set SearchRange = Tbl.ListColumns("Datum").Range
    Set Result = SearchRange.Find(What:=SelectedDate, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No PlanMealFood found for Date {0}", SelectedDate)
        Exit Sub
    End If
    
    FirstFoundAddress = Result.Address
    
    Do
        MealId = Tbl.ListColumns("MahlzeitenId").Range.Rows(Result.Row).Value
        If MealId = Me.Id Then
            Set Food = New NutritionPlanMealFood
            
            Food.FoodId = Tbl.ListColumns("NahrungsmittelId").Range(Result.Row).Value
            Food.PlanDate = Tbl.ListColumns("Datum").Range(Result.Row).Value
            Food.MealId = Tbl.ListColumns("MahlzeitenId").Range(Result.Row).Value
            Food.Name = Tbl.ListColumns("Nahrungsmittel").Range(Result.Row).Value
            Food.Brand = Tbl.ListColumns("Marke").Range(Result.Row).Value
            Food.Unit = Tbl.ListColumns("Einheit").Range(Result.Row).Value
            Food.Amount = Tbl.ListColumns("Menge").Range(Result.Row).Value
            Food.Nutrients.Calories = Tbl.ListColumns("Kalorien").Range(Result.Row).Value
            Food.Nutrients.Macros.Protein = Tbl.ListColumns("Proteine").Range(Result.Row).Value
            Food.Nutrients.Macros.Carbohydrates = Tbl.ListColumns("Kohlenhydrate").Range(Result.Row).Value
            Food.Nutrients.Macros.Sugar = Tbl.ListColumns("Zucker").Range(Result.Row).Value
            Food.Nutrients.Macros.Fat = Tbl.ListColumns("Fett").Range(Result.Row).Value
            Food.Nutrients.Micros.Fiber = Tbl.ListColumns("Ballaststoffe").Range(Result.Row).Value

            Me.TryAddFood Food.FoodId, Food
        End If
        
        Set Result = SearchRange.FindNext(Result)
    Loop While Not Result Is Nothing And FirstFoundAddress <> Result.Address
End Sub
Public Function TryAddFood(FoodId As Integer, Food As NutritionPlanMealFood) As Boolean

    If Me.Foods.Exists(FoodId) Then
        TryAddFood = False
        Exit Function
    End If
    
    Me.AddFood FoodId, Food
    TryAddFood = True
End Function

Public Sub AddFood(FoodId As Integer, Food As NutritionPlanMealFood)
    Me.Foods.Add FoodId, Food
End Sub

Public Sub Delete()
    '@TODO: Muss noch entfernt werden aus den NutritionPlan Object
    NutritionPlanDatabase.DeleteMeal Me
End Sub


Public Function GetNutrients() As Nutrients
    Dim NewNutrients As New Nutrients
    Dim Key As Variant
    For Each Key In Me.Foods.Keys
        NewNutrients.Calories = NewNutrients.Calories + Foods(Key).Nutrients.Calories
        NewNutrients.Macros.Protein = NewNutrients.Macros.Protein + Foods(Key).Nutrients.Macros.Protein
        NewNutrients.Macros.Carbohydrates = NewNutrients.Macros.Carbohydrates + Foods(Key).Nutrients.Macros.Carbohydrates
        NewNutrients.Macros.Sugar = NewNutrients.Macros.Sugar + Foods(Key).Nutrients.Macros.Sugar
        NewNutrients.Macros.Fat = NewNutrients.Macros.Fat + Foods(Key).Nutrients.Macros.Fat
        
        NewNutrients.Micros.Fiber = NewNutrients.Micros.Fiber + Foods(Key).Nutrients.Micros.Fiber
    Next Key
    Set GetNutrients = NewNutrients
End Function
Public Function GetButton() As Shape
    Dim Shp As Shape, ShpDup As Shape
    Dim MealNutrients As New Nutrients
    
    Set MealNutrients = Me.GetNutrients
    
    Set Shp = Worksheets("Templates").Shapes("BtnTemplate_PlanMeal")
    
    Set ShpDup = Shp.Duplicate

    ShpDup.GroupItems("BtnTemplate_PlanMeal_Meal").TextEffect.Text = Printf("Mahlzeit {0}", Me.Id)
    ShpDup.GroupItems("BtnTemplate_PlanMeal_Foods").TextEffect.Text = Printf("Nahrungsmittel: {0} Stück", Me.Foods.Count)
    ShpDup.GroupItems("BtnTemplate_PlanMeal_Macros").TextEffect.Text = Printf("{0} Kcal. - (P: {1} | K: {2} | F: {3})", MealNutrients.Calories, MealNutrients.Macros.Protein, MealNutrients.Macros.Carbohydrates, MealNutrients.Macros.Fat)
    
    
    ShpDup.GroupItems("BtnTemplate_PlanMeal_Meal").Name = Printf("BtnPlanMeal_Meal_{0}", Me.Id)
    ShpDup.GroupItems("BtnTemplate_PlanMeal_Foods").Name = Printf("BtnPlanMeal_Foods_{0}", Me.Id)
    ShpDup.GroupItems("BtnTemplate_PlanMeal_Macros").Name = Printf("BtnPlanMeal_Macros_{0}", Me.Id)
    
    ShpDup.GroupItems("BtnTemplate_PlanMeal_Delete").Name = Printf("BtnPlanMeal_Delete_{0}", Me.Id)
    ShpDup.Name = Printf("BtnPlanMeal_{0}", Me.Id)
    
    ShpDup.OnAction = "'PlanMealButton_Click """ & Me.Id & """'"
    ShpDup.GroupItems(Printf("BtnPlanMeal_Delete_{0}", Me.Id)).OnAction = "'PlanMealButton_Delete_Click """ & Me.Id & """'"
    
    Set GetButton = ShpDup
End Function
