Attribute VB_Name = "NutritionConfigs"
Public Const NutritionPlanMealsWorksheetName = "Rohdaten_PlanMahlzeit"
Public Const NutritionPlanMealsListObjectName = "TblMeals"

Public Const NutritionPlanMealFoodsWorksheetName = "Rohdaten_MahlzeitLebensmittel"
Public Const NutritionPlanMealFoodsListObjectName = "TblMealFoods"

Property Get MealTable() As ListObject
    Set MealTable = Functions.GetListObject(NutritionPlanMealsListObjectName, NutritionPlanMealsWorksheetName)
End Property

Property Get MealFoodTable() As ListObject
    Set MealFoodTable = Functions.GetListObject(NutritionPlanMealFoodsListObjectName, NutritionPlanMealFoodsWorksheetName)
End Property
