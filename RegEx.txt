Function ExtractVT(InputText As String) As String
    Dim RegEx As Object
    Dim Match As Object
    
    ' Create RegEx object
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .Pattern = "V\d{4}\sT\d{4}"   ' Look for V#### T####
        .Global = False
        .IgnoreCase = True
    End With
    
    If RegEx.Test(InputText) Then
        Set Match = RegEx.Execute(InputText)(0)
        ExtractVT = Match.Value
    Else
        ExtractVT = "" ' return blank if not found
    End If
End Function


Sub TestExtract()
    Dim s As String
    s = "Invoice number is V1234 T5678 for customer"
    MsgBox ExtractVT(s)   ' Output: V1234 T5678
End Sub
