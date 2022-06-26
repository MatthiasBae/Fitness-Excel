VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NutritionPlanMealFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FoodId As Integer
Public PlanDate As Date
Public MealId As Integer
Public Name As String
Public Brand As String
Public Unit As String
Public Amount As Double
Public Nutrients As New Nutrients

Public Sub Delete()
    NutritionPlanDatabase.DeleteMealFood Me
End Sub
Public Function GetButton() As Shape
    Dim Shp As Shape, ShpDup As Shape
    Dim MealNutrients As New Nutrients

    Set Shp = Worksheets("Templates").Shapes("BtnTemplate_MealFood")
    
    Set ShpDup = Shp.Duplicate

    ShpDup.GroupItems("BtnTemplate_MealFood_Food").TextEffect.Text = Printf("{0}", Me.Name)
    ShpDup.GroupItems("BtnTemplate_MealFood_Brand").TextEffect.Text = Printf("{0}", Me.Brand)
    ShpDup.GroupItems("BtnTemplate_MealFood_Macros").TextEffect.Text = Printf("{0} Kcal. - (P: {1} | K: {2} | F: {3})", Me.Nutrients.Calories, Me.Nutrients.Macros.Protein, Me.Nutrients.Macros.Carbohydrates, Me.Nutrients.Macros.Fat)
    
    
    ShpDup.GroupItems("BtnTemplate_MealFood_Food").Name = Printf("BtnPlanFood_Food_{0}", Me.FoodId)
    ShpDup.GroupItems("BtnTemplate_MealFood_Brand").Name = Printf("BtnPlanFood_Brand_{0}", Me.FoodId)
    ShpDup.GroupItems("BtnTemplate_MealFood_Macros").Name = Printf("BtnPlanFood_Macros_{0}", Me.FoodId)
    ShpDup.GroupItems("BtnTemplate_MealFood_Delete").OnAction = "'PleanMealFoodButton_Delete_Click """ & Me.FoodId & """'"
    ShpDup.GroupItems("BtnTemplate_MealFood_Delete").Name = Printf("BtnPlanFood_Delete_{0}", Me.FoodId)
    ShpDup.Name = Printf("BtnPlanFood_{0}", Me.FoodId)
    
    Set GetButton = ShpDup
End Function
