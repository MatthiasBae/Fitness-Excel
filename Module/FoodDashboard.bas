Attribute VB_Name = "FoodDashboard"
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
Private Function pSelectedFoodUnit(Optional Item As FoodUnit = Nothing) As FoodUnit
    Static CurrentItem As FoodUnit
    If Not Item Is Nothing Then
        Set CurrentItem = Item
        Exit Function
    End If
    
    Set pSelectedFoodUnit = CurrentItem
End Function

Public Property Get SelectedFoodUnit() As FoodUnit
    Set SelectedFoodUnit = pSelectedFoodUnit
End Property
Public Property Set SelectedFoodUnit(Item As FoodUnit)
    pSelectedFoodUnit Item
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
Public Sub Init()

End Sub

Public Sub Reset()
    FoodDashboard.ResetFoodList
    FoodDashboard.ResetSelectedFoodPanel
    FoodDashboard.ResetSelectedFoodUnitPanel
End Sub

Public Function PrepareFoodList(SelectedRng As Range, Name As String, Brand As String, Optional TopCount As Integer) As WrapPanel
    Dim WrapPanel As New WrapPanel
    Dim FoodList As New Dictionary, FoodId As Variant, SelectedFood As Food
    Dim FoodBtn As Shape
    
    Set FoodList = FoodDatabase.GetFoods(Name, Brand, TopCount)
    
    WrapPanel.Initialize SelectedRng, 1
    
    For Each FoodId In FoodList.Keys
        Set SelectedFood = FoodList(FoodId)
        Set FoodBtn = SelectedFood.GetButton
        
        WrapPanel.Add FoodBtn
    Next
    Set PrepareFoodList = WrapPanel
End Function
Public Sub FillFoodList()
    Dim Ws As Worksheet
    Dim Name As String, Brand As String, TopCount As Integer
    
    Set Ws = Worksheets("Dashboard Lebensmittel")
    
    Name = Ws.Range("Text_Fd_SearchFood")
    Brand = Ws.Range("Text_Fd_SearchBrand")
    TopCount = Ws.Range("Text_Fd_SearchTop")
    
    'ResetFoodList
    FoodList.Clear
    
    Set FoodList = FoodDashboard.PrepareFoodList(Ws.Range("List_Fd_FoodEntries"), Name, Brand, TopCount)
    FoodList.Render
    
    Application.CutCopyMode = False
End Sub

Public Sub ResetFoodList()
    Dim Shp As Shape
    For Each Shp In Worksheets("Dashboard Lebensmittel").Shapes
        If InStr(1, Shp.Name, "BtnFood") Then
            Shp.Delete
        End If
    Next Shp
End Sub
Public Sub FillSelectedFoodPanel(Item As Food)
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Dashboard Lebensmittel")
    
    Dim DefaultUnit As New FoodUnit
    Set DefaultUnit = Item.GetDefaultUnit()
    
    Ws.Range("Text_Fd_FoodSelectedName").Value = Item.Name
    Ws.Range("Text_Fd_FoodSelectedBrand").Value = Item.Brand
    Ws.Range("Text_Fd_SelectedFoodUnitAmount").Value = DefaultUnit.Amount
    Ws.Range("List_Fd_FoodSelectedUnits").Value = DefaultUnit.Name
    Ws.Range("List_Fd_ACG1").Value = Item.ACG1
    Ws.Range("List_Fd_ACG2").Value = Item.ACG2
    Ws.Range("List_Fd_ACG3").Value = Item.ACG3
    With Ws.Range("List_Fd_FoodSelectedUnits").Validation
        .Delete
        .Add Type:=xlValidateList, Operator:= _
        xlBetween, Formula1:=Item.GetUnitNames
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
        .InputTitle = ""
        .InputMessage = ""
        .ShowInput = True
    End With

End Sub

Public Sub FillSelectedFoodUnitPanel(Item As FoodUnit)
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Dashboard Lebensmittel")
    
    Ws.Range("Text_Fd_SelectedFoodUnitAmount").Value = Item.Amount
    Ws.Range("Text_Fd_SelectedFoodUnitCalories").Value = Item.Nutrients.Calories
    Ws.Range("Text_Fd_SelectedFoodUnitProtein").Value = Item.Nutrients.Macros.Protein
    Ws.Range("Text_Fd_SelectedFoodUnitCarbs").Value = Item.Nutrients.Macros.Carbohydrates
    Ws.Range("Text_Fd_SelectedFoodUnitSugar").Value = Item.Nutrients.Macros.Sugar
    Ws.Range("Text_Fd_SelectedFoodUnitFat").Value = Item.Nutrients.Macros.Fat
End Sub


Public Sub ResetSelectedFoodPanel()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Dashboard Lebensmittel")
    
    Ws.Range("Text_Fd_FoodSelectedName").Value = ""
    Ws.Range("Text_Fd_FoodSelectedBrand").Value = ""
    Ws.Range("Text_Fd_SelectedFoodUnitAmount").Value = 0
    Ws.Range("List_Fd_FoodSelectedUnits").Value = ""
    Ws.Range("List_Fd_ACG1").Value = ""
    Ws.Range("List_Fd_ACG2").Value = ""
    Ws.Range("List_Fd_ACG3").Value = ""
    With Ws.Range("List_Fd_FoodSelectedUnits").Validation
        .Delete
    End With
End Sub

Public Sub ResetSelectedFoodUnitPanel()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Dashboard Lebensmittel")
    
    Ws.Range("Text_Fd_SelectedFoodUnitCalories").Value = 0
    Ws.Range("Text_Fd_SelectedFoodUnitProtein").Value = 0
    Ws.Range("Text_Fd_SelectedFoodUnitCarbs").Value = 0
    Ws.Range("Text_Fd_SelectedFoodUnitSugar").Value = 0
    Ws.Range("Text_Fd_SelectedFoodUnitFat").Value = 0
End Sub

Public Sub SaveFood()
    Dim Name As String, Brand As String, Unit As String, Amount As Double, ACG1 As String, ACG2 As String, ACG3 As String
    Dim Calories As Double, Protein As Double, Carbs As Double, Sugar As Double, Fat As Double
    Dim Ws As Worksheet
    
    Set Ws = ThisWorkbook.Worksheets("Dashboard Lebensmittel")
    
    Name = Ws.Range("Text_Fd_FoodSelectedName").Value
    Brand = Ws.Range("Text_Fd_FoodSelectedBrand").Value
    Unit = Ws.Range("List_Fd_FoodSelectedUnits").Value
    Amount = Ws.Range("Text_Fd_SelectedFoodUnitAmount").Value
    ACG1 = Ws.Range("List_Fd_ACG1").Value
    ACG2 = Ws.Range("List_Fd_ACG2").Value
    ACG3 = Ws.Range("List_Fd_ACG3").Value
    
    Calories = Ws.Range("Text_Fd_SelectedFoodUnitCalories").Value
    Protein = Ws.Range("Text_Fd_SelectedFoodUnitProtein").Value
    Carbs = Ws.Range("Text_Fd_SelectedFoodUnitCarbs").Value
    Sugar = Ws.Range("Text_Fd_SelectedFoodUnitSugar").Value
    Fat = Ws.Range("Text_Fd_SelectedFoodUnitFat").Value
    
    If Name = "" Or Unit = "" Or Amount <= 0 Then
        MsgBox "Bitte alle Informationen angeben", vbExclamation, "Datenbank"
        Exit Sub
    End If
    
    FoodDatabase.SaveFood Name, Brand, Unit, Amount, Calories, Protein, Carbs, Sugar, Fat, ACG1, ACG2, ACG3
    MsgBox Printf("{0} {1} wurde gespeichert", Brand, Name), vbInformation, "Datenbank"
End Sub

Public Sub FillCategoryLists()

End Sub
