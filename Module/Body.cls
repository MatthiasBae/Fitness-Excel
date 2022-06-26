VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public PlanDate As Date
Public Weight As Double
Public Fat As Double
Public BMI As Double

Public Sub Load(SelectedDate As Date)
    Dim Tbl As ListObject
    Dim Result As Range

    Set Tbl = BodyConfigs.BodyTable
    Set Result = Tbl.ListColumns("Datum").Range.Find(What:=SelectedDate, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No Body found with for date {0}", SelectedDate)
        Exit Sub
    End If
    
    Me.PlanDate = Tbl.ListColumns("Datum").Range.Rows(Result.Row).Value
    Me.Weight = Round(Tbl.ListColumns("Gewicht").Range.Rows(Result.Row).Value, 1)
    Me.Fat = Round(Tbl.ListColumns("Fett").Range.Rows(Result.Row).Value, 1)
    Me.BMI = Round(Tbl.ListColumns("BMI").Range.Rows(Result.Row).Value, 1)
End Sub

Public Sub Delete()
    Dim Tbl As ListObject
    Dim Result As Range

    Set Tbl = BodyConfigs.BodyTable
    Set Result = Tbl.ListColumns("Datum").Range.Find(What:=Me.PlanDate, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Result Is Nothing Then
        Debug.Print Printf("No Body found with for date {0}", SelectedDate)
        Exit Sub
        
    End If
    
    Result.EntireRow.Delete
End Sub

Public Function GetButton() As Shape
    '@TODO: einbauen von kleinen Pfeilen welche die Veränderung anzeigen
    Dim Shp As Shape, ShpDup As Shape

    Set Shp = Worksheets("Templates").Shapes("BtnTemplate_Body")
    
    Set ShpDup = Shp.Duplicate

    ShpDup.GroupItems("BtnTemplate_Body_Date").TextEffect.Text = Printf("{0}", Me.PlanDate)
    ShpDup.GroupItems("BtnTemplate_Body_Weight").TextEffect.Text = Printf("Gewicht: {0} Kg", Me.Weight)
    ShpDup.GroupItems("BtnTemplate_Body_Fat").TextEffect.Text = Printf("Fett: {0}%", Me.Fat)
    
    ShpDup.GroupItems("BtnTemplate_Body_Date").Name = Printf("BtnBody_Date_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    ShpDup.GroupItems("BtnTemplate_Body_Weight").Name = Printf("BtnBody_Weight_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    ShpDup.GroupItems("BtnTemplate_Body_Fat").Name = Printf("BtnBody_Fat_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    ShpDup.GroupItems("BtnTemplate_Body_Delete").OnAction = "DeleteBody"
    ShpDup.GroupItems("BtnTemplate_Body_Delete").Name = Printf("BtnBody_Delete_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    ShpDup.Name = Printf("BtnBody_{0}", Format(Me.PlanDate, "yyyyMMdd"))
    
    Set GetButton = ShpDup
End Function

