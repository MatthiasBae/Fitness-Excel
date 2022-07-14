Attribute VB_Name = "RecipesDashboard"
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
    RecipesDashboard.ResetRecipeList
    'RecipeDashboard.ResetSelectedFoodPanel
    'RecipeDashboard.ResetSelectedFoodUnitPanel
End Sub

Public Function PrepareRecipeList(SelectedRng As Range, Name As String, Brand As String, Optional TopCount As Integer) As WrapPanel
    Dim WrapPanel As New WrapPanel
    Dim FoodList As New Dictionary, FoodId As Variant, SelectedFood As Food
    Dim FoodBtn As Shape
    
    Set FoodList = FoodDatabase.GetFoods(Name, Brand, TopCount, True)
    
    WrapPanel.Initialize SelectedRng, 1
    
    For Each FoodId In FoodList.Keys
        Set SelectedFood = FoodList(FoodId)
        Set FoodBtn = SelectedFood.GetButton
        
        WrapPanel.Add FoodBtn
    Next
    Set PrepareRecipeList = WrapPanel
End Function
Public Sub FillRecipeList()
    Dim Ws As Worksheet
    Dim Name As String, Brand As String, TopCount As Integer
    
    Set Ws = Worksheets("Dashboard Rezepte")
    
    Name = Ws.Range("Text_Rc_SearchRecipe")
    TopCount = Ws.Range("Text_Rc_SearchTop")
    
    'ResetFoodList
    FoodList.Clear
    
    Set FoodList = RecipesDashboard.PrepareRecipeList(Ws.Range("List_Rc_RecipeEntries"), Name, Brand, TopCount)
    FoodList.Render
    
    Application.CutCopyMode = False
End Sub

Public Sub ResetRecipeList()
    Dim Shp As Shape
    For Each Shp In Worksheets("Dashboard Rezepte").Shapes
        If InStr(1, Shp.Name, "BtnFood") Then
            Shp.Delete
        End If
    Next Shp
End Sub
