Option Explicit
'UPDATED BY: IVAN Y - CI

Private Enum Clean
    NumCheckBox = 50
End Enum
Sub RenameCheckBox()
    Selection.Name = "Check Box 50"
End Sub
Sub CheckBoxSelect()
    Dim cws As Worksheet
    Dim Boxcheck As Boolean
    Dim y As Integer, col As Byte
    Dim dict_default As Object
    Set cws = ThisWorkbook.Worksheets("Clean Macro")
    'DETERMINE THE CHECKBOX STATE BASED ON THE VALUE IN CELL(M3)
    With cws
        Select Case .Cells(3, 13).Value
            Case "Uncheck All"
                col = 0
                Boxcheck = False
            Case "Check All"
                col = 0
                Boxcheck = True
            Case "Default - MFC - Entity Level"
                col = 9
            Case "Default - MFC - Class"
                col = 10
            Case "Default - MFT/UT"
                col = 11
            Case Else
                col = 0 'DEFAULT CASE IF NONE MATCH
        End Select
    'LOOP THROUGH CHECKBOXES AND SET THEIR VALUES BASED ON THE DETERMINED COLUMN
    For y = 1 To NumCheckBox Step 1
        If col > 0 Then
            If .Cells(y + 1, col).Value = "ON" Then
                .CheckBoxes("Check Box " & y).Value = True
            Else
                .CheckBoxes("Check Box " & y).Value = False
            End If
        ElseIf col = 0 Then
            .CheckBoxes("Check Box " & y).Value = Boxcheck
        End If
    Next y
    End With
End Sub
Sub CheckClean()
    ' DECLARE VARIABLES
    Dim wb As Workbook
    Dim ws As Worksheet, cws As Worksheet
    Dim link_Export As String, link_Clean As String
    Dim lastRow As Long
    Dim fileName As String, FundCode As String
    Dim wsTabs As Variant
    Dim Dict_wsTabs As Object, Dict_wsCheck As Object
    Dim y As Integer, x As Integer
    Dim foundCell As Range
    
    'TURN OFF EXCEL APPLICATION SETTINGS FOR FASTER EXECUTION.
    Call TurnOffApp
    
    Set cws = ThisWorkbook.Worksheets("Clean Macro")
    
    ' CREATE A DICTIONARY OBJECT
    Set Dict_wsTabs = CreateObject("Scripting.Dictionary")
    With cws
        lastRow = .Cells(.Rows.Count, "P").End(xlUp).Row
        .Range("O5:Q" & lastRow + 10).Clear
        For y = 1 To NumCheckBox Step 1
            If .CheckBoxes("Check Box " & y).Value = 1 Then
                Dict_wsTabs.Add .Cells(y + 1, 8).Value, ""
            End If
        Next y
        link_Export = .Cells(2, "P").Value & "\"
    End With
    
    y = 5
    fileName = Dir(link_Export & "*.xlsx", vbReadOnly = True)
    'LOOP THROUGH EACH CLEAN FILES IN THE DIRECTORY
    Do While fileName <> ""
        Set Dict_wsCheck = CreateObject("Scripting.Dictionary")
        Set Dict_wsCheck = CloneDict(Dict_wsTabs)
        With cws
            .Cells(y, "P").Value = fileName
            .Cells(y, "O").Value = Trim(Split(fileName, "-")(1))
        End With
        Set wb = Workbooks.Open(link_Export & fileName, vbReadOnly = True)
        
        'LOOP THROUGH EACH WORKSHEET IN THE WORKBOOK
        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVeryHidden Then
            Else
                If Dict_wsCheck.Exists(ws.Name) Then
                    Dict_wsCheck.Remove (ws.Name)
                    'CHECK FOR INVESTIGATE SECURITY IN THE "SUSPENDED LOSS CONTINUITY" SHEET
                    If ws.Name = "Suspended Loss Continuity" Then
                        Set foundCell = ws.columns("C").Find(What:="Investigate Security", LookIn:=xlValues, LookAt:=xlPart)
                        If Not foundCell Is Nothing Then
                            cws.Cells(y, "Q").Value = "[Investigate Security] Exist"
                            Exit For
                        End If
                    End If
                Else
                    cws.Cells(y, "Q").Value = "Unnecessary Tab - [" & ws.Name & "] Exist"
                    Exit For
                End If
            End If
        Next ws
        
        'DETERMINE THE STATUS OF THE CHECKS
        If cws.Cells(y, "Q").Value = "" Then
            If Dict_wsCheck.Count > 0 Then
                cws.Cells(y, "Q").Value = "Missing Tab; " & Dict_wsCheck.Keys()(0)
            Else
                cws.Cells(y, "Q").Value = "Looks Good."
            End If
        End If
        
        wb.Close False
        y = y + 1
        fileName = Dir
    Loop
    
    Call TurnOnApp
    MsgBox "Completed"

End Sub
'CLONE THE KEYS FROM ORIGINAL DICTIONARY
Function CloneDict(Dict As Object) As Object
  Dim newDict
  Dim key As Variant
  
  Set newDict = CreateObject("Scripting.Dictionary")
  For Each key In Dict.Keys
    newDict.Add key, ""
  Next
  
  newDict.CompareMode = Dict.CompareMode

  Set CloneDict = newDict
End Function

Sub CreateClean()
    ' Declare variables
    Dim wb As Workbook, cwb As Workbook
    Dim ws As Worksheet, cws As Worksheet
    Dim link_Export As String, link_Clean As String
    Dim fileName As String, FundCode As String
    Dim wsTabs As Variant
    Dim Dict_wsTabs As Object
    Dim y As Integer

    
    ' Set cwb to the current workbook and cws to the "Clean Macro" worksheet
    Set cwb = ThisWorkbook
    Set cws = cwb.Worksheets("Clean Macro")
    
    ' Create a dictionary object
    Set Dict_wsTabs = CreateObject("Scripting.Dictionary")
    With cws
        y = .Cells(Rows.Count, 3).End(xlUp).Row
        ' Clear a specific range in the "Clean Macro" worksheet
        .Range(.Cells(5, 2), .Cells(y + 10, 4)).Clear
        
        ' Check the value of each checkbox in the "Clean Macro" worksheet
        ' If a checkbox is checked (value = 1), add the corresponding value from column 8 to the dictionary
        For y = 1 To NumCheckBox Step 1
            If .CheckBoxes("Check Box " & y).Value = 1 Then
                Dict_wsTabs.Add .Cells(y + 1, 8).Value, ""
            Else
            End If
        Next y
        'RETRIEVE THE EXPORT AND CLEAN FOLDERS PATHS FROM CELLS C1 and C2.
        link_Export = .Cells(1, 3).Value & "\"
        link_Clean = .Cells(2, 3).Value & "\"
    End With
    
    Dim lastRow As Integer
    Dim Dict_SLC As Object
    Set Dict_SLC = CreateObject("Scripting.Dictionary")
    If Dict_wsTabs.Exists("Suspended Loss Continuity") = True Then
        Set ws = cwb.Worksheets("Lists")
        With ws
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            For y = 3 To lastRow Step 1
                If Dict_SLC.Exists(.Cells(y, 1).Value) = False Then
                    Dict_SLC.Add .Cells(y, 1).Value, .Cells(y, 2).Value
                End If
            Next y
        End With
    End If
    ' Initialize y to 5
    y = 5
    ' Get the name of the first read-only Excel file in the export folder
    fileName = Dir(link_Export & "*.xlsx", vbReadOnly = True)
    ' Enter a loop that continues as long as Filename is not an empty string
    ' This would mean there are no more files to process
    Dim Str_SID As String
    Dim x As Integer
    Dim foundCell As Range
    Do While fileName <> ""
    ' Write the filename and the FundCode to the "Clean Macro" worksheet
        With cws
            .Cells(y, 3).Value = fileName
            .Cells(y, 2).Value = Trim(Split(fileName, "-")(1))
        End With
        ' Open the workbook in read-only mode
        Set wb = Workbooks.Open(link_Export & fileName, vbReadOnly = True)
        
        'SUSPENDED LOSS CONTINUITY - Investigate Security  (Investigate Security)
        If Dict_wsTabs.Exists("Suspended Loss Continuity") And WorksheetExists("Suspended Loss Continuity", wb) Then
            Set ws = wb.Worksheets("Suspended Loss Continuity")
            With ws
                lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
                For x = 2 To lastRow Step 1
                    Set foundCell = .Range("C" & x & ":C" & lastRow).Find(What:="Investigate Security", LookIn:=xlValues, LookAt:=xlPart)
                    If Not foundCell Is Nothing Then
                        x = foundCell.Row
                        Str_SID = .Cells(x, "D").Value
                        If Dict_SLC.Exists(Str_SID) = True Then
                            .Cells(x, "C").Value = Dict_SLC(Str_SID)
                            .Cells(x, "E").Value = Str_SID
                        Else
                            cws.Cells(y, "D").Value = "[Unknown Investigate Security] - " & Str_SID
                        End If
                    Else
                        Exit For
                    End If
                Next x
            End With
        End If
        
        ' Loop through each worksheet in the workbook
        ' If the worksheet name exists in the dictionary or the worksheet is not visible, change the tab color to blue and delete the worksheet
        ' Otherwise, just delete the worksheet
        For Each ws In wb.Worksheets
            If Dict_wsTabs.Exists(ws.Name) Or ws.Visible <> True Then
                ws.Tab.ThemeColor = xlThemeColorAccent1
                ws.Tab.TintAndShade = -0.249977111117893 'changes the tab colour to bluews.Delete
            Else
                ws.Delete
            End If
            
        Next ws
        'BREAK ALL LINKS AND SAVE THE CLEANED WORKBOOK IN THE CLEAN FOLDER WITH THE SAME FILENAME AND CLOSE IT.
        Call BreakAllLink(wb)
        
        wb.SaveAs fileName:=link_Clean & fileName & ".xlsx"
        
        wb.Close
        ' Increment y by 1 and get the next filename
        y = y + 1
        
        fileName = Dir
    Loop
    
    ' Turn the Excel application settings back on and display a message box indicating the successful completion of the operation
    Call TurnOnApp
    MsgBox "Completed"

End Sub

Sub BreakAllLink(wb As Workbook)
    Dim Arr_Links As Variant
    Dim y As Byte
    Arr_Links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    'CHECK IF THERE ARE ANY LINKS TO BREAK
        If Not IsEmpty(Arr_Links) Then
        'LOOP THROUGH EACH LINK AND BREAK IT
        For y = LBound(Arr_Links) To UBound(Arr_Links)
            wb.BreakLink Name:=Arr_Links(y), Type:=xlLinkTypeExcelLinks
        Next y
    End If
End Sub

Function WorksheetExists(wsName As String, wb As Workbook) As Boolean
    On Error Resume Next
    WorksheetExists = (wb.Worksheets(wsName).Name = wsName)
End Function

Public Sub TurnOffApp()
    With Application
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .AskToUpdateLinks = False
    End With
End Sub
Public Sub TurnOnApp()
    With Application
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .AskToUpdateLinks = True
    End With
End Sub
