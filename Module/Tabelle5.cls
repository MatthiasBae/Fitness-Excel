VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = Target.Worksheet.Range("Text_St_CaloriesFormulaText").Address Then
        SettingsDashboard.CreateFormula
    ElseIf Target.Address = Target.Worksheet.Range("List_St_FormulaTypes").Address Then
        SettingsDashboard.LoadCaloriesFormula
    
    End If
End Sub
