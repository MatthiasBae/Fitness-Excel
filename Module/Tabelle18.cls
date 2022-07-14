VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim UnitName As String
    
    If Target.Address = Target.Worksheet.Range("List_Fd_FoodSelectedUnits").Address Then
        UnitName = Target.Worksheet.Range("List_Fd_FoodSelectedUnits").Value
        If UnitName = "" Then
            Exit Sub
        End If
        
        If FoodDatabase.FoodUnitExists(FoodDashboard.SelectedFood.FoodId, UnitName) Is Nothing Then
            Exit Sub
        End If
        
        Set FoodDashboard.SelectedFoodUnit = FoodDashboard.SelectedFood.GetUnit(UnitName)
        FoodDashboard.FillSelectedFoodUnitPanel FoodDashboard.SelectedFoodUnit
    End If
End Sub

