VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WrapPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ListObjects As New Dictionary

Private Columns As Long
Private Section As Range

Private Sub Class_Initialize()
    Colums = 1
End Sub

Public Sub Initialize(rng As Range, columnCount As Long)
    Me.SetColumns columnCount
    Set Section = rng
End Sub

Public Sub Add(Btn As Shape)
    Dim Shp As Shape
    Dim Name As String
    
    If Section Is Nothing Then
        Exit Sub
    End If
    Name = Btn.Name
    
    Btn.Cut
    Section.Worksheet.Paste
    Set Btn = Section.Worksheet.Shapes(Name)
    Btn.Visible = msoFalse
    ListObjects.Add Btn.Name, Btn
End Sub
Public Sub Remove(Key As Variant)
    ListObjects.Remove Key
End Sub
Public Function List() As Dictionary
    Set List = ListObjects
End Function

Public Sub SetColumns(Optional columnCount As Long = 0)
    If columnCount = 0 Then
        Columns = 1
        Exit Sub
    End If
    
    Columns = columnCount
End Sub

Public Sub Render()
    Dim Clm As Long, Rw As Long
    Dim ipc As Long, itemCounter As Long
    
    Dim Key As Long, Item As Shape
    Dim x, y As Long
    x = Section.Left
    y = Section.Top
    
    ipc = ItemsPerColumn
    currentIpc = 1
    itemCounter = 0
    For Rw = 1 To ipc
        For Clm = 1 To Columns
        
            Set Item = ListObjects.Items(itemCounter)
            
            y = (Item.Height * (Rw - 1)) + Section.Top
            x = (Item.Width * (Clm - 1)) + Section.Left
            
            Item.Left = x
            Item.Top = y
            Item.Visible = msoTrue
            
            itemCounter = itemCounter + 1
        Next Clm
    Next Rw
    Cells(1, 1).Select
End Sub
Public Sub Clear()
    Dim Item As Shape, Key As Variant
    For Each Key In ListObjects.Keys
        Set Item = ListObjects(Key)
        Item.Delete
    Next Key
    ListObjects.RemoveAll
End Sub

Private Function ItemsPerColumn()
    ItemsPerColumn = Round(ListObjects.Count / Columns, 0)
End Function
