Attribute VB_Name = "Functions"
Public Function Printf(ByVal strText As String, ParamArray args()) As String
    Dim i As Long
    For i = 0 To UBound(args)
        strText = Replace$(strText, "{" & i & "}", args(i))
    Next
    Printf = strText
End Function

Public Function DateDiff(DateFrom As Date, DateTo As Date) As Integer
    DateDiff = DateTo - DateFrom
End Function

Public Function DateAdd(DateFrom As Date, Days As Integer) As Date
    DateAdd = DateFrom + Days
End Function

Public Function GetIdFromButton(ButtonName As String) As String
    Dim i As Integer
    For i = Len(ButtonName) To 0 Step -1
        If Mid(ButtonName, i, 1) = "_" Then
            GetIdFromButton = Mid(ButtonName, i + 1, Len(ButtonName) - i + 1)
            Exit For
        End If
    Next i
End Function

Public Function GetListObject(ListObjectName As String, WorksheetName As String) As ListObject
    Dim Ws As Worksheet
    
    Set Ws = ThisWorkbook.Worksheets(WorksheetName)
    Set GetListObject = Ws.ListObjects(ListObjectName)
End Function
Public Function SortDictionaryByKey(dict As Object, Optional sortorder As XlSortOrder = xlAscending) As Object
    
    Dim arrList As Object
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    ' Put keys in an ArrayList
    Dim Key As Variant, coll As New Collection
    For Each Key In dict
        arrList.Add Key
    Next Key
    
    ' Sort the keys
    arrList.Sort
    
    ' For descending order, reverse
    If sortorder = xlDescending Then
        arrList.Reverse
    End If
    
    ' Create new dictionary
    Dim dictNew As Object
    Set dictNew = CreateObject("Scripting.Dictionary")
    
    ' Read through the sorted keys and add to new dictionary
    For Each Key In arrList
        dictNew.Add Key, dict(Key)
    Next Key
    
    ' Clean up
    Set arrList = Nothing
    Set dict = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByKey = dictNew
        
End Function

