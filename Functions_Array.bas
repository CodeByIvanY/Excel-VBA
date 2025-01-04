Public Function BubbleSrt(ArrayIn, Ascending As Boolean)
Dim SrtTemp As Variant
Dim i As Long
Dim j As Long
If Ascending = True Then
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) > ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
Else
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) < ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
End If
BubbleSrt = ArrayIn
End Function

Array2Sort = BubbleSrt(Array2Sort, True)
'True being sort as Ascending. False will sort Decending
'***********************************************************************************************
'BREAK ALL LINKS
Sub BreakAllLinks()
    Dim Arr_Links As Variant
    Dim wb As Workbook
    Dim y As Byte
    'SET CURRENT WORKBOOK AND GET ALL EXCEL LINKS IN THE WORKBOOK
    Set wb = ThisWorkbook
    Arr_Links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    'CHECK IF THERE ARE ANY LINKS TO BREAK
    If Not IsEmpty(Arr_Links) Then
    'LOOP THROUGH EACH LINK AND BREAK IT
    For y = LBound(Arr_Links) To UBound(Arr_Links)
        wb.BreakLink Name:=Arr_Links(y), Type:=xlLinkTypeExcelLinks
    Next y
    End If
End Sub

