Attribute VB_Name = "BodyConfigs"
Public Const BodyWorksheetName = "Rohdaten_Körper"
Public Const BodyListObjectName = "TblBody"

Property Get BodyTable() As ListObject
    Set BodyTable = Functions.GetListObject(BodyListObjectName, BodyWorksheetName)
End Property

