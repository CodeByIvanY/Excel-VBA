Function HighlightNonZeroAndErrors(rng As Range)
    ' Clear existing conditional formats
    rng.FormatConditions.Delete

    ' Highlight cells not equal to 0
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=0")
        .SetFirstPriority
        .Font.Color = RGB(192, 0, 0) ' DARK RED
        .Interior.Color = RGB(255, 199, 206) ' LIGHT RED FILL
        .StopIfTrue = False
    End With

    ' Highlight cells with errors
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISERROR(" & rng.Cells(1, 1).Address(False, False) & ")")
        .Font.Color = RGB(192, 0, 0) ' DARK RED
        .Interior.Color = RGB(255, 199, 206) ' LIGHT RED FILL
        .StopIfTrue = False
    End With
End Function

Function HighlightDuplicates(rng As Range)
    ' CONDITIONAL FORMATTING: HIGHLIGHT DUPLICATES IN COLUMN A - FUND CODE
    With rng
        .FormatConditions.Delete
        With .FormatConditions.AddUniqueValues
            .DupeUnique = xlDuplicate
            .Font.Color = RGB(192, 0, 0) ' DARK RED
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.Color = RGB(255, 199, 206) ' LIGHT RED FILL
            .StopIfTrue = False
        End With
    End With
End Function

Function HighlightFlagText(rng As Range, Optional keyword As String = "FLAG")
    ' CONDITIONAL FORMATTING: HIGHLIGHT CELLS CONTAINING "FLAG"
    With rng
        .FormatConditions.Delete
        With .FormatConditions.Add(Type:=xlTextString, String:=keyword, TextOperator:=xlContains)
            .Font.Color = RGB(192, 0, 0) ' DARK RED TEXT
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.Color = RGB(255, 199, 206) ' LIGHT RED FILL
            .StopIfTrue = False
        End With
    End With
End Function

Function HighlightGreaterABSOnesAndErrors(rng As Range)
    rng.FormatConditions.Delete
    Dim fillColor As Long: fillColor = RGB(255, 199, 206) ' Light Red Fill
    Dim fontColor As Long: fontColor = RGB(192, 0, 0)     ' Dark Red Text

    ' Highlight cells greater than 1
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="1")
        .SetFirstPriority
        .Font.Color = fontColor
        .Interior.Color = fillColor
        .StopIfTrue = False
    End With

    ' Highlight cells less than -1
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="-1")
        .Font.Color = fontColor
        .Interior.Color = fillColor
        .StopIfTrue = False
    End With

    ' Highlight cells with errors
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISERROR(" & rng.Cells(1, 1).Address(False, False) & ")")
        .Font.Color = fontColor
        .Interior.Color = fillColor
        .StopIfTrue = False
    End With
End Function

Sub ApplyHeaderFormatting(rng As Range, _
                          Optional headerData As Variant, _
                          Optional InteriorColor As Long = -1, _
                          Optional FontColor As Long = -1, _
                          Optional HorizontalAlignment As XlHAlign = xlLeft)

    ' APPLY HEADER FORMATTING WITH OPTIONAL HEADER VALUES AND COLOR
    With rng
        ' SSET HEADER TEXT IF PROIVIDED
        If Not IsMissing(headerData) Then
            .Value = headerData
        End If
        
        ' APPLY FILL COLOR DEFAULT - GOLD
        If InteriorColor = -1 Then
            .Interior.Color = RGB(255, 192, 0)
        Else
            .Interior.Color = InteriorColor
        End If
        
        If FontColor <> -1 Then
        Else
            .Font.Color = RGB(255, 192, 0)
        End If
        ' APPLY FONT AND DEFAULT
        .Font.Bold = True
        .HorizontalAlignment = HorizontalAlignment
    End With
End Sub






