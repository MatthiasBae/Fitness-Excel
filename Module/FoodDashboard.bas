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
    
    Ws.Range("Text_Fd_FoodSelectedName").Value = Item.Name
    Ws.Range("Text_Fd_FoodSelectedBrand").Value = Item.Brand
    Ws.Range("Text_Fd_SelectedFoodUnitAmount").Value = Item.GetDefaultUnit().Amount
    Ws.Range("List_Fd_FoodSelectedUnits").Value = Item.GetDefaultUnit().Name
    With Ws.Range("List_Fd_FoodSelectedUnits").Validation
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
    Set Ws = ThisWorkbook.Worksheets("Dashboard Lebensmittel")
    
    Ws.Range("Text_Fd_FoodSelectedName").Value = ""
    Ws.Range("Text_Fd_FoodSelectedBrand").Value = ""
    Ws.Range("Text_Fd_SelectedFoodUnitAmount").Value = 0
    Ws.Range("List_Fd_FoodSelectedUnits").Value = ""
    With Ws.Range("List_Fd_FoodSelectedUnits").Validation
        .Delete
    End With
End Sub

