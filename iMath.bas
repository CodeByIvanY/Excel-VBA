'iMax accepts a variable number of arguments using ParamArray and returns the largest numeric value.
Function iMax(ParamArray args() As Variant) as Double
    Dim x As Long
    iMax = args(0)
    For x = 1 To Ubound(args)
        If args(x) > iMax Then
            iMax = args(x)
        End If
    Next X
End Function

Function iMin(ParamArray args() As Variant) as Double
    Dim x As Long
    iMax = args(0)
    For x = 1 To Ubound(args)
        If args(x) < iMax Then
            iMax = args(x)
        End If
    Next X
End Function


Function iMaxArr(arr() as Double) as Double
    Dim x As Byte
    iMax = arr(0)
    For x = 1 To Ubound(arr)
        If arr(x) > iMax Then
            iMax = arr(x)
        End If
    Next X
End Function


'iSum is to calculate the total of values in a specified worksheet column, starting from a designated row down to the last populated row in that column.
Function iSum(cws as Worksheet, iCol as String, iStartRow As Byte) as Double
    Dim lastRow as Long, y as Long
    iSum = 0
    with cws
        On Error Resume Next
        .ShowAllData
        On Error Goto 0
        lastRow = .Cells(.Rows.Count, iCol).End(XlUp).Row
        If lastRow < iStartRow Then Exit Function
        
        For y = iStartRow To lastRow Step 1
            iSum = iSum + .Cells(y, iCol).value
        Next y
    End with
End Function
