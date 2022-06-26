VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FoodUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public Amount As Double
Public IsDefault As Boolean
Public Nutrients As New Nutrients

Public Function ReCalculate(ToAmount As Double) As Nutrients
    Dim CalcNutrients As New Nutrients
    
    CalcNutrients.Calories = Nutrients.Calories / Amount * ToAmount
    CalcNutrients.Macros.Protein = Nutrients.Macros.Protein / Amount * ToAmount
    CalcNutrients.Macros.Carbohydrates = Nutrients.Macros.Carbohydrates / Amount * ToAmount
    CalcNutrients.Macros.Sugar = Nutrients.Macros.Sugar / Amount * ToAmount
    CalcNutrients.Macros.Fat = Nutrients.Macros.Fat / Amount * ToAmount
    'CalcNutrients.Macros.SaturatedFat = Nutrients.Macros.SaturatedFat / Amount * ToAmount
    
    CalcNutrients.Micros.Fiber = Nutrients.Micros.Fiber / Amount * ToAmount
    Set ReCalculate = CalcNutrients
End Function
