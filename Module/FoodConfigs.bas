Attribute VB_Name = "FoodConfigs"
Public Const FoodsWorksheetName = "Rohdaten_Lebensmittel"
Public Const FoodsListObjectName = "TblFoods"

Public Const FoodUnitsWorksheetName = "Rohdaten_LebensmittelEinheiten"
Public Const FoodUnitsListObjectName = "TblFoodUnits"

Public Const FoodIngredientsWorksheetName = "Rohdaten_LebensmittelZutaten"
Public Const FoodIngredientsListObjectName = "TblFoodIngredients"

Property Get FoodTable() As ListObject
    Set FoodTable = Functions.GetListObject(FoodsListObjectName, FoodsWorksheetName)
End Property

Property Get FoodUnitsTable() As ListObject
    Set FoodUnitsTable = Functions.GetListObject(FoodUnitsListObjectName, FoodUnitsWorksheetName)
End Property

Property Get FoodIngredientsTable() As ListObject
    Set FoodIngredientsTable = Functions.GetListObject(FoodIngredientsListObjectName, FoodIngredientsWorksheetName)
End Property
