VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Food"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FoodId As Long
Public CompId As Long
Public Store As String
Public Name As String
Public Brand As String
Public ACG1 As String
Public ACG2 As String
Public ACG3 As String

Public Attributes As New FoodAttributes
Public Units As New Dictionary

Public Sub Load(SelectedId As Long)
    Me.LoadFood SelectedId
    Me.LoadFoodUnits SelectedId
End Sub

Public Sub LoadFood(SelectedId As Long)
    Dim Ws As Worksheet, Tbl As ListObject
    Dim Result As Range
    
    Set Tbl = FoodConfigs.FoodTable
    Set Result = Tbl.ListColumns("NahrungsmittelId").Range.Find(What:=SelectedId, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No Food found with Id {0}", SelectedId)
        Exit Sub
    End If
    
    Me.FoodId = Tbl.ListColumns("NahrungsmittelId").Range.Rows(Result.Row).Value
    Me.CompId = Tbl.ListColumns("VergleichsId").Range.Rows(Result.Row).Value
    Me.Store = Tbl.ListColumns("Laden").Range.Rows(Result.Row).Value
    Me.Name = Tbl.ListColumns("Nahrungsmittel").Range.Rows(Result.Row).Value
    Me.Brand = Tbl.ListColumns("Marke").Range.Rows(Result.Row).Value
    Me.ACG1 = Tbl.ListColumns("KategorieLevel1").Range.Rows(Result.Row).Value
    Me.ACG2 = Tbl.ListColumns("KategorieLevel2").Range.Rows(Result.Row).Value
    Me.ACG3 = Tbl.ListColumns("KategorieLevel3").Range.Rows(Result.Row).Value
    Attributes.Gluten = Tbl.ListColumns("Gluten").Range.Rows(Result.Row).Value
    Attributes.NutritionScore = Tbl.ListColumns("Nutritionscore").Range.Rows(Result.Row).Value
    Attributes.Vegan = Tbl.ListColumns("Vegan").Range.Rows(Result.Row).Value
    Attributes.Vegetarian = Tbl.ListColumns("Vegetarisch").Range.Rows(Result.Row).Value

End Sub

Public Sub LoadFoodUnits(SelectedId As Long)
    Dim Tbl As ListObject
    Dim Result As Range, SearchRange As Range, FirstFoundAddress As String
    
    Dim Unit As FoodUnit

    Set Tbl = FoodConfigs.FoodUnitsTable
    Set SearchRange = Tbl.ListColumns("NahrungsmittelId").Range
    Set Result = SearchRange.Find(What:=SelectedId, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No FoodUnit found for Food found with Id {0}", SelectedId)
        Exit Sub
    End If
    
    FirstFoundAddress = Result.Address
    
    Do
        Set Unit = New FoodUnit
        
        Unit.Amount = Tbl.ListColumns("Menge").Range.Rows(Result.Row).Value
        Unit.IsDefault = Tbl.ListColumns("Standardeinheit").Range.Rows(Result.Row).Value
        Unit.Name = Tbl.ListColumns("Einheit").Range.Rows(Result.Row).Value
        
        Unit.Nutrients.Calories = Tbl.ListColumns("Kalorien").Range.Rows(Result.Row).Value
        Unit.Nutrients.Macros.Protein = Tbl.ListColumns("Proteine").Range.Rows(Result.Row).Value
        Unit.Nutrients.Macros.Carbohydrates = Tbl.ListColumns("Kohlenhydrate").Range.Rows(Result.Row).Value
        Unit.Nutrients.Macros.Sugar = Tbl.ListColumns("Zucker").Range.Rows(Result.Row).Value
        Unit.Nutrients.Macros.Fat = Tbl.ListColumns("Fett").Range.Rows(Result.Row).Value
        
        Unit.Nutrients.Micros.Fiber = Tbl.ListColumns("Ballaststoffe").Range.Rows(Result.Row).Value
        
        Me.TryAddUnit Unit.Name, Unit
        
        Set Result = SearchRange.FindNext(Result)
    Loop While Not Result Is Nothing And FirstFoundAddress <> Result.Address
End Sub


Public Function TryAddUnit(Name As String, Unit As FoodUnit) As Boolean

    If Me.Units.Exists(Name) Then
        TryAddUnit = False
        Exit Function
    End If
    
    Me.AddUnit Name, Unit
    TryAddUnit = True
End Function

Public Sub AddUnit(Name As String, Unit As FoodUnit)
    Me.Units.Add Name, Unit
End Sub

Public Function GetDefaultUnit() As FoodUnit
    Dim Unit As New FoodUnit
    Dim Item As New FoodUnit, Key As Variant
    
    For Each Key In Me.Units.Keys
        Set Unit = Me.Units(Key)
        If Unit.IsDefault = True Then
            Set GetDefaultUnit = Unit
            Exit Function
        End If
    Next Key
    
    Set GetDefaultUnit = Me.Units.Items(0)
        
End Function

Public Function GetUnitNames() As String
    Dim Key As Variant
    Dim UnitText As String
    
    For Each Key In Me.Units.Keys
        UnitText = UnitText & Key & ","
    Next Key
    
    GetUnitNames = Left(UnitText, Len(UnitText) - 1)
End Function

Public Function GetUnit(UnitName As String) As FoodUnit
    Set GetUnit = Me.Units(UnitName)
End Function
Public Function GetButton() As Shape
    Dim Shp As Shape, ShpDup As Shape
    Dim Unit As New FoodUnit
    
    Set Shp = Worksheets("Templates").Shapes("BtnTemplate_Food")
    
    Set ShpDup = Shp.Duplicate
    
    Set Unit = Me.GetDefaultUnit
    
    ShpDup.GroupItems("BtnTemplate_Food_Foodname").TextEffect.Text = Printf("{0}", Me.Name)
    ShpDup.GroupItems("BtnTemplate_Food_Brandname").TextEffect.Text = Printf("{0}", Me.Brand)
    ShpDup.GroupItems("BtnTemplate_Food_Macros").TextEffect.Text = Printf("Auf {0} {1} {2} Kcal. - (P: {3} | K: {4} | F: {5})", Unit.Amount, Unit.Name, Unit.Nutrients.Calories, Unit.Nutrients.Macros.Protein, Unit.Nutrients.Macros.Carbohydrates, Unit.Nutrients.Macros.Fat)
    
    ShpDup.GroupItems("BtnTemplate_Food_Foodname").Name = Printf("BtnFood_Foodname_{0}", Me.FoodId)
    ShpDup.GroupItems("BtnTemplate_Food_Brandname").Name = Printf("BtnFood_Brandname_{0}", Me.FoodId)
    ShpDup.GroupItems("BtnTemplate_Food_Macros").Name = Printf("BtnFood_Macros_{0}", Me.FoodId)
    ShpDup.GroupItems("BtnTemplate_Food_Background").Name = Printf("BtnFood_Background_{0}", Me.FoodId)
    ShpDup.Name = Printf("BtnFood_{0}", Me.FoodId)
    'ShpDup.OnAction = "FoodButton_Click"
    ShpDup.OnAction = "'FoodButton_Click """ & Me.FoodId & """'"
    
    Set GetButton = ShpDup
End Function


