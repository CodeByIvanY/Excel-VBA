'CLONE THE KEYS FROM ORIGINAL DICTIONARY
Function CloneDict(Dict As Object) As Object
  Dim newDict
  Dim key As Variant
  
  Set newDict = CreateObject("Scripting.Dictionary")
  For Each key In Dict.Keys
    newDict.Add key, Dict(key)
  Next
  
  newDict.CompareMode = Dict.CompareMode

  Set CloneDict = newDict
End Function

Dim Dict_New As Object
Set Dict_New = CreateObject("Scripting.Dictionary")
Set Dict_New = CloneDict(Dict_Old)

'-----------------------------------------------------------------------------------------------------------------------'
Function SortDictKey(dict As Object, Optional sortorder As XlSortOrder = xlAscending) As Object
    Dim arrList As Object
    Set arrList = CreateObject("System.Collections.Arraylist")
    Dim key As Variant, coll As New Collection
    For Each key In dict
        arrList.Add key
    Next key
    arrList.Sort
    Dim dictNew As Object
    Set dictNew = CreateObject("Scripting.Dictionary")
    For Each key In arrList
        dictNew.Add key, dict(key)
    Next key
    Set arrList = Nothing
    Set dict = Nothing
    Set SortDictKey = dictNew
End Function
'-----------------------------------------------------------------------------------------------------------------------'
Function F_DictList(wsName As String, Col_Key As String, Col_Item As String, StartRow As Integer, Col_LastRow As String, Col_Key2 As String, Col_Item2 As String) As Object
    Dim ws As Worksheet
    Dim lastRow As Long, y As Long
    Dim Dict As Object
    Dim Str_Key As String, Str_Item As String
    Set ws = ThisWorkbook.Worksheets(wsName)
    Set Dict = CreateObject("Scripting.Dictionary")
    With ws
        If .FilterMode = True Then
        .ShowAllData
        End If
        lastRow = .Cells(Rows.Count, Col_LastRow).End(xlUp).Row
        For y = StartRow To lastRow Step 1
            If Col_Key2 = "0" Then
                Str_Key = .Cells(y, Col_Key).Value
            Else
                Str_Key = .Cells(y, Col_Key).Value & " - " & .Cells(y, Col_Key2).Value
            End If
            
            If Dict.Exists(Str_Key) Then
            Else
                If Col_Item2 = "0" Then
                    Str_Item = .Cells(y, Col_Item).Value
                Else
                    Str_Item = .Cells(y, Col_Item).Value & " - " & .Cells(y, Col_Item2).Value
                End If
                Dict.Add Str_Key, Str_Item
            End If
        Next y
    End With
    Set F_DictList = Dict
End Function

