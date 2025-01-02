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
