Attribute VB_Name = "SettingsDashboard"

Public Sub CreateFormula()
    Dim Ws As Worksheet
    Dim FormulaTextRange As Range, FormulaTextValuesRange As Range, FormulaValueRange As Range
    Dim FormulaTextValues As String, KPIName As String, KPIValue As String
    
    Dim AvailableKPIRange As Range, KPI As Range
    
    On Error Resume Next
    
    Set Ws = Worksheets("Einstellungen")

    Set FormulaTextRange = Ws.Range("TextCaloriesFormulaText")
    Set FormulaTextValuesRange = Ws.Range("TextCaloriesFormulaTextValues")
    Set FormulaValueRange = Ws.Range("TextCaloriesFormulaValue")
    
    Set AvailableKPIRange = Worksheets("Rohdaten_KPIs").Range("A2", Worksheets("Rohdaten_KPIs").Range("A2").End(xlDown))
    
    FormulaTextValues = FormulaTextRange.Value
    For Each KPI In AvailableKPIRange
        KPIName = KPI.Value
        KPIValue = KPI.Worksheet.Cells(KPI.Row, 3)
        FormulaTextValues = Replace(FormulaTextValues, KPIName, KPIValue)
    Next KPI
    FormulaTextValuesRange.Value = FormulaTextValues
    FormulaValueRange.FormulaLocal = "=" & FormulaTextValues
End Sub

Public Sub AddKPI()
    Dim Ws As Worksheet
    Dim FormulaTextRange As Range, KPIListRange As Range
    
    Set Ws = Worksheets("Einstellungen")

    Set FormulaTextRange = Ws.Range("TextCaloriesFormulaText")
    Set KPIListRange = Ws.Range("ListKPIs")
    FormulaTextRange.Value = FormulaTextRange.Value & KPIListRange.Value
End Sub

Public Sub LoadCaloriesFormula()
    Dim Ws As Worksheet, WsRawdata As Worksheet
    Dim FormulaTypeListRange As Range, FormulaSourceListRange As Range
    
    Dim i As Integer
    
    Set Ws = Worksheets("Einstellungen")
    Set FormulaTypeListRange = Ws.Range("ListFormulaTypes")
    Set FormulaSourceListRange = Ws.Range("ListCaloriesFormulaSource")
    
    Set WsRawdata = Worksheets("Rohdaten_Kalorienformeln")
    
    i = 1
    Do Until WsRawdata.Cells(i, 1) = ""
        
        If WsRawdata.Cells(i, 1) = FormulaSourceListRange.Value _
            And WsRawdata.Cells(i, 2) = FormulaTypeListRange.Value Then
            
            Ws.Range("TextCaloriesFormulaText") = WsRawdata.Cells(i, 3)
            Exit Do
        End If
        
        i = i + 1
    Loop
End Sub

Public Sub SaveCaloriesFormula()
    Dim Ws As Worksheet, WsRawdata As Worksheet
    Dim FormulaTypeListRange As Range, FormulaSourceListRange As Range
    
    Dim i As Integer
    
    Set Ws = Worksheets("Einstellungen")
    Set FormulaTypeListRange = Ws.Range("ListFormulaTypes")
    Set FormulaSourceListRange = Ws.Range("ListCaloriesFormulaSource")
    
    Set WsRawdata = Worksheets("Rohdaten_Kalorienformeln")
    i = 1
    Do Until WsRawdata.Cells(i, 1) = ""
        
        If WsRawdata.Cells(i, 1) = FormulaSourceListRange.Value _
            And WsRawdata.Cells(i, 2) = FormulaTypeListRange.Value Then
            
            WsRawdata.Cells(i, 3) = Ws.Range("TextCaloriesFormulaText")
            WsRawdata.Cells(i, 4) = Ws.Range("TextCaloriesFormulaTextValues")
            WsRawdata.Cells(i, 5).FormulaLocal = Ws.Range("TextCaloriesFormulaValue").FormulaLocal
            Exit Do
        End If
        
        i = i + 1
    Loop
    
End Sub
