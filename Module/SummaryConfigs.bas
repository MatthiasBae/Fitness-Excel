Attribute VB_Name = "SummaryConfigs"
Public Const SummaryWorksheetName = "Rohdaten_Zusammenfassung"
Public Const SummaryListObjectName = "TblSummary"

Property Get SummaryTable() As ListObject
    Set SummaryTable = Functions.GetListObject(SummaryListObjectName, SummaryWorksheetName)
End Property
