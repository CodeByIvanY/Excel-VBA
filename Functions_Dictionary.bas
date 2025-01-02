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

'-------------------------------------------------------------------------
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
