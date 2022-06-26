Attribute VB_Name = "FoodDashboard"
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
