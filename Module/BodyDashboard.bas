Attribute VB_Name = "BodyDashboard"
'@Singleton
Private Function pSelectedBody(Optional Item As Body = Nothing) As Body
    Static CurrentItem As Body
    If Not Item Is Nothing Then
        Set CurrentItem = Item
        Exit Function
    End If
    
    Set pSelectedBody = CurrentItem
End Function

Public Property Get SelectedBody() As Body
    Set SelectedBody = pSelectedBody
End Property
Public Property Set SelectedBody(Item As Body)
    pSelectedBody Item
End Property

'@Singleton
Public Function pBodyList(Optional List As WrapPanel = Nothing) As WrapPanel
    Static CurrentList As WrapPanel
    If Not List Is Nothing Then
        Set CurrentList = List
        Exit Function
    End If
    
    If CurrentList Is Nothing Then
        Set CurrentList = New WrapPanel
    End If
    
    Set pBodyList = CurrentList
End Function
Public Property Get BodyList() As WrapPanel
    Set BodyList = pBodyList
End Property
Public Property Set BodyList(Item As WrapPanel)
    pBodyList Item
End Property
Public Sub Init()

End Sub
Public Sub Reset()
    BodyDashboard.ResetBodyList
End Sub
Public Function PrepareBodyList(SelectedRng As Range, DateFrom As Date, Optional WeightFilter As String = "", Optional FatFilter As String = "") As WrapPanel
    Dim WrapPanel As New WrapPanel
    Dim BodyList As New Dictionary, BodyId As Variant, SelectedBody As Body
    Dim BodyBtn As Shape
    
    Set BodyList = BodyDatabase.GetBodies(DateFrom, WeightFilter, FatFilter)
    
    WrapPanel.Initialize SelectedRng, 1

    For Each BodyId In BodyList.Keys
        Set SelectedBody = BodyList(BodyId)
        Set BodyBtn = SelectedBody.GetButton
        
        WrapPanel.Add BodyBtn
    Next
    Set PrepareBodyList = WrapPanel
End Function

Public Sub FillBodyList()
    Dim Ws As Worksheet
    Dim DateFrom As Date, WeightFilter As String, FatFilter As String
    
    Set Ws = Worksheets("Dashboard Körper")

    DateFrom = IIf(Ws.Range("Text_Bd_SearchDateFrom") = "", Date, Ws.Range("Text_Bd_SearchDateFrom"))
    WeightFilter = Ws.Range("Text_Bd_SearchWeight")
    FatFilter = Ws.Range("Text_Bd_SearchFat")
    
    BodyDashboard.BodyList.Clear
    Set BodyDashboard.BodyList = BodyDashboard.PrepareBodyList(Ws.Range("List_Bd_BodyEntries"), DateFrom, WeightFilter, FatFilter)
    BodyDashboard.BodyList.Render
End Sub

Public Sub ResetBodyList()
    Dim Shp As Shape
    For Each Shp In Worksheets("Dashboard Körper").Shapes
        If InStr(1, Shp.Name, "BtnBody_") Then
            Shp.Delete
        End If
    Next Shp
End Sub

