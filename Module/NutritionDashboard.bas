Attribute VB_Name = "NutritionDashboard"
'@Singleton
Private Function pSelectedFood(Optional Item As Food = Nothing) As Food
    Static CurrentItem As Food
    If Not Item Is Nothing Then
        Set CurrentItem = Item
        Exit Function
    End If
    
    Set pSelectedFood = CurrentItem
End Function

Public Property Get SelectedFood() As Food
    Set SelectedFood = pSelectedFood
End Property
Public Property Set SelectedFood(Item As Food)
    pSelectedFood Item
End Property

'@Singleton
Private Function pSelectedPlan(Optional Item As NutritionPlan = Nothing) As NutritionPlan
    Static CurrentItem As NutritionPlan
    If Not Item Is Nothing Then
        Set CurrentItem = Item
        Exit Function
    End If
    
    Set pSelectedPlan = CurrentItem
End Function
Public Property Get SelectedPlan() As NutritionPlan
    Set SelectedPlan = pSelectedPlan
End Property
Public Property Set SelectedPlan(Item As NutritionPlan)
    pSelectedPlan Item
End Property

'@Singleton
Private Function pSelectedPlanMeal(Optional Item As NutritionPlanMeal = Nothing) As NutritionPlanMeal
    Static CurrentItem As NutritionPlanMeal
    If Not Item Is Nothing Then
        Set CurrentItem = Item
        Exit Function
    End If
    
    Set pSelectedPlanMeal = CurrentItem
End Function
Public Property Get SelectedPlanMeal() As NutritionPlanMeal
    Set SelectedPlanMeal = pSelectedPlanMeal
End Property
Public Property Set SelectedPlanMeal(Item As NutritionPlanMeal)
    pSelectedPlanMeal Item
End Property

'@Singleton
Public Function pFoodList(Optional List As WrapPanel = Nothing) As WrapPanel
    Static CurrentList As WrapPanel
    If Not List Is Nothing Then
        Set CurrentList = List
        Exit Function
    End If
    
    If CurrentList Is Nothing Then
        Set CurrentList = New WrapPanel
    End If
    
    Set pFoodList = CurrentList
End Function
Public Property Get FoodList() As WrapPanel
    Set FoodList = pFoodList
End Property
Public Property Set FoodList(Item As WrapPanel)
    pFoodList Item
End Property

'@Singleton
Public Function pPlanList(Optional List As WrapPanel = Nothing) As WrapPanel
    Static CurrentList As WrapPanel
    If Not List Is Nothing Then
        Set CurrentList = List
        Exit Function
    End If
    
    If CurrentList Is Nothing Then
        Set CurrentList = New WrapPanel
    End If
    
    Set pPlanList = CurrentList
End Function
Public Property Get PlanList() As WrapPanel
    Set PlanList = pPlanList
End Property
Public Property Set PlanList(Item As WrapPanel)
    pPlanList Item
End Property

Public Sub FoodButton_Click(FoodId As Long)
    Dim FoodItem As New Food
    FoodItem.Load FoodId
    
    Set NutritionDashboard.SelectedFood = FoodItem
    NutritionDashboard.FillSelectedFoodPanel NutritionDashboard.SelectedFood
    
End Sub

Public Sub PlanMealButton_Click(MealId As Integer)
    Dim MealItem As New NutritionPlanMeal
    
    If NutritionDashboard.SelectedPlan Is Nothing Then
        Set NutritionDashboard.SelectedPlan = New NutritionPlan
        NutritionDashboard.SelectedPlan.Load Worksheets("Dashboard Ernährung").Range("TextDateFrom")
    End If
    
    Set MealItem = NutritionDashboard.SelectedPlan.Meals(MealId)
    
    Set NutritionDashboard.SelectedPlanMeal = MealItem
    NutritionDashboard.FillPlanMealFoodList NutritionDashboard.SelectedPlanMeal
End Sub

Public Sub PlanMealButton_Delete_Click(MealId As Integer)
    Dim MealItem As New NutritionPlanMeal
    
    If NutritionDashboard.SelectedPlan Is Nothing Then
        Set NutritionDashboard.SelectedPlan = New NutritionPlan
        NutritionDashboard.SelectedPlan.Load Worksheets("Dashboard Ernährung").Range("TextDateFrom")
    End If
    

    Set MealItem = NutritionDashboard.SelectedPlan.Meals(MealId)
    MealItem.Delete
    NutritionDashboard.SelectedPlan.Meals.Remove MealItem.Id
    NutritionDashboard.FillPlanMealList NutritionDashboard.SelectedPlan
End Sub
Public Sub PleanMealFoodButton_Delete_Click(FoodId As Long)
    Dim MealFoodItem As New NutritionPlanMealFood
    
    If NutritionDashboard.SelectedPlan Is Nothing Then
        Set NutritionDashboard.SelectedPlan = New NutritionPlan
        NutritionDashboard.SelectedPlan.Load Worksheets("Dashboard Ernährung").Range("TextDateFrom")
    End If
    
    If NutritionDashboard.SelectedPlanMeal Is Nothing Then
        Set NutritionDashboard.SelectedPlanMeal = NutritionDashboard.SelectedPlan.Meals(Worksheets("Dashboard Ernährung").Range("TextMealNr"))
    End If
    
    Set MealFoodItem = NutritionDashboard.SelectedPlanMeal.Foods(FoodId)
    
    If NutritionDashboard.SelectedPlanMeal.Foods.Count > 1 Then
        MealFoodItem.Delete
        NutritionDashboard.SelectedPlanMeal.Foods.Remove FoodId
    Else
        NutritionDashboard.SelectedPlanMeal.Delete
        NutritionDashboard.SelectedPlan.Meals.Remove MealFoodItem.MealId
    End If
    NutritionDashboard.FillPlanMealFoodList NutritionDashboard.SelectedPlanMeal
End Sub

Public Sub Init()
    Dim ActualDate As Date
    ActualDate = Date
    
    Worksheets("Dashboard Ernährung").Range("TextDateFrom") = ActualDate
    Worksheets("Dashboard Ernährung").Range("TextDateTo") = ActualDate + 7
End Sub
Public Sub reset()
    NutritionDashboard.ResetSelectedFoodPanel
    
    NutritionDashboard.ResetFoodList
    NutritionDashboard.ResetPlanList
End Sub
Public Sub FillSelectedFoodPanel(Item As Food)
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Dashboard Ernährung")
    
    Ws.Range("TextFoodSelectedName").Value = Item.Name
    Ws.Range("TextFoodSelectedBrand").Value = Item.Brand
    Ws.Range("TextFoodSelectedAmount").Value = Item.GetDefaultUnit().Amount
    Ws.Range("ListFoodSelectedUnits").Value = Item.GetDefaultUnit().Name
    With Ws.Range("ListFoodSelectedUnits").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Item.GetUnitNames
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub
Public Sub ResetSelectedFoodPanel()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Dashboard Ernährung")
    
    Ws.Range("TextFoodSelectedName").Value = ""
    Ws.Range("TextFoodSelectedBrand").Value = ""
    Ws.Range("TextFoodSelectedAmount").Value = 0
    Ws.Range("ListFoodSelectedUnits").Value = ""
    With Ws.Range("ListFoodSelectedUnits").Validation
        .Delete
    End With
End Sub

Public Function PrepareFoodList(SelectedRng As Range, Name As String, Brand As String, Optional TopCount As Integer) As WrapPanel
    Dim WrapPanel As New WrapPanel
    Dim FoodList As New Dictionary, FoodId As Variant, SelectedFood As Food
    Dim FoodBtn As Shape
    
    Set FoodList = FoodDatabase.GetFoods(Name, Brand, TopCount)
    
    WrapPanel.Initialize SelectedRng, 1
    
    Application.ScreenUpdating = False
    For Each FoodId In FoodList.Keys
        Set FoodBtn = FoodList(FoodId).GetButton
        
        WrapPanel.Add FoodBtn
    Next
    Set PrepareFoodList = WrapPanel
    Application.ScreenUpdating = True
End Function
Public Sub FillFoodList()
    Dim Ws As Worksheet
    Dim Name As String, Brand As String, TopCount As Integer
    
    Set Ws = Worksheets("Dashboard Ernährung")
    
    Name = Ws.Range("TextSearchFoodField")
    Brand = Ws.Range("TextSearchBrandField")
    TopCount = Ws.Range("TextSearchTopField")
    
    'ResetFoodList
    FoodList.Clear
    
    Set FoodList = NutritionDashboard.PrepareFoodList(Ws.Range("ListFoods"), Name, Brand, TopCount)
    FoodList.Render
    
    Application.CutCopyMode = False
End Sub

Public Sub ResetFoodList()
    Dim Shp As Shape
    For Each Shp In Worksheets("Dashboard Ernährung").Shapes
        If InStr(1, Shp.Name, "BtnFood") Then
            Shp.Delete
        End If
    Next Shp
End Sub
Public Sub ResetPlanList()
    Dim Shp As Shape
    For Each Shp In Worksheets("Dashboard Ernährung").Shapes
        If InStr(1, Shp.Name, "BtnPlan") Then
            Shp.Delete
        End If
    Next Shp
End Sub
Public Function PreparePlanMealList(SelectedRng As Range, Plan As NutritionPlan) As WrapPanel
    Dim WrapPanel As New WrapPanel
    Dim PlanMealList As New Dictionary, Item As Variant
    Dim PlanMealBtn As Shape
    
    Set PlanMealList = Plan.Meals
    
    WrapPanel.Initialize SelectedRng, 1

    For Each Item In PlanMealList.Keys
        If PlanMealList(Item) Is Nothing Then
            Exit For
        End If
        Set PlanMealBtn = PlanMealList(Item).GetButton
        
        WrapPanel.Add PlanMealBtn
    Next
    Set PreparePlanMealList = WrapPanel
End Function
Public Sub FillPlanMealList(Plan As NutritionPlan)
    Dim Ws As Worksheet
    Set Ws = Worksheets("Dashboard Ernährung")

    PlanList.Clear
    Set PlanList = NutritionDashboard.PreparePlanMealList(Ws.Range("ListPlans"), Plan)
    PlanList.Render
End Sub

Public Function PreparePlanMealFoodList(SelectedRng As Range, PlanMeal As NutritionPlanMeal) As WrapPanel
    Dim WrapPanel As New WrapPanel
    Dim PlanMealFoodList As New Dictionary, Item As Variant
    Dim PlanMealFoodBtn As Shape

    Set PlanMealFoodList = PlanMeal.Foods
    
    WrapPanel.Initialize SelectedRng, 1
    
    For Each Item In PlanMealFoodList.Keys
        If PlanMealFoodList(Item) Is Nothing Then
            Exit For
        End If
        Set PlanMealFoodBtn = PlanMealFoodList(Item).GetButton
        
        WrapPanel.Add PlanMealFoodBtn
    Next
    Set PreparePlanMealFoodList = WrapPanel
End Function

Public Sub FillPlanMealFoodList(PlanMeal As NutritionPlanMeal)
    Dim Ws As Worksheet
    Set Ws = Worksheets("Dashboard Ernährung")
    
    PlanList.Clear
    Set PlanList = NutritionDashboard.PreparePlanMealFoodList(Ws.Range("ListPlans"), PlanMeal)
    PlanList.Render
End Sub


Public Sub AddFoodToPlan()
    Dim Ws As Worksheet
    Set Ws = Worksheets("Dashboard Ernährung")

    Dim DateFrom As Date, DateTo As Date, IsCheatMeal As Boolean, Weekday As Integer, MealId As Integer, Amount As Double, Unit As String
    
    DateFrom = Ws.Range("TextDateFrom").Value
    DateTo = Ws.Range("TextDateTo").Value
    IsCheatMeal = IIf(Ws.Range("BoolIsCheatmeal").Value = "Ja", True, False)
    Weekday = IIf(Ws.Range("ListWeekday").Value = "", 0, Ws.Range("ListWeekday").Value)
    MealId = IIf(Ws.Range("TextMealNr").Value <= 0, 1, Ws.Range("TextMealNr").Value)
    Amount = Ws.Range("TextFoodSelectedAmount").Value
    Unit = IIf(Ws.Range("ListFoodSelectedUnits").Value = "", "Gramm", Ws.Range("ListFoodSelectedUnits").Value)

    NutritionPlanDatabase.TryAddFood NutritionDashboard.SelectedFood, Unit, Amount, MealId, DateFrom, DateTo, IsCheatMeal, Weekday
    NutritionDashboard.SelectedPlan.Load DateFrom
    FillPlanMealList NutritionDashboard.SelectedPlan
End Sub
