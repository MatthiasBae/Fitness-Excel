VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = Target.Worksheet.Range("TextDateFrom").Address Then
        Set NutritionDashboard.SelectedPlan = New NutritionPlan
        NutritionDashboard.SelectedPlan.Load Target.Worksheet.Range("TextDateFrom").Value
        
        NutritionDashboard.FillPlanMealList NutritionDashboard.SelectedPlan
    End If
End Sub

