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
