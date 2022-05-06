' Timesheet Helper Comments
'
' 3.20 - 3 January 2021 - Added ShName for Flex 410, changed Work Schedule Selected to TS Helper, other functions not used
'
'
'***************************************************
Option Explicit

'File filter for Excel workbooks
Private Const WorkbookFileFilter = "Excel Workbooks (*.xls; *.xlsb; *.xlsm; *.xlsx),*.xls;*.xlsb;*.xlsm;*.xlsx"

'Worksheet Names
Public Const Instructions_ShName = "Instructions"
Public Const Labor_Flex980_ShName = "Labor_Flex980"
Public Const Labor_Flex980_2weeks_ShName = "Labor_Flex980_2weeks"
Public Const Labor_Flex410_ShName = "Labor_Flex410"                    ' Added TSHelper 3.20
Public Const Simple_Labor_Adjust_ShName = "Simple Labor Adjustment"
Public Const Dropdown_Entries_ShName = "Dropdown_Entries"

'First and Last Row of entries for each Labor worksheet
Public Const FirstLaborRow_Flex980 = 10
Public Const LastLaborRow_Flex980 = 289
Public Const FirstLaborRow_Flex980_2weeks = 10
Public Const LastLaborRow_Flex980_2weeks = 289
Public Const FirstLaborRow_Flex410 = 10                                 ' Added TSHelper 3.20
Public Const LastLaborRow_Flex410 = 289                                 ' Added TSHelper 3.20

'Work Schedule list entries
Public Const WorkSchedule_Flex980 = "Flex 9/80"
Public Const WorkSchedule_Flex980_2weeks = "Flex 9/80 (two-week view)"
Public Const WorkSchedule_Flex410 = "Flex 4/10"                        ' Added TSHelper 3.20

'Selected work schedule
Public WorkSchedule

'Workbook and worksheet variables used by import routines
Private DataBookName
Private ThisBookName
Private ThisSheetName

Sub Get_Selected_Work_Schedule()
    'Get selected Work Schedule
    WorkSchedule = Range("WorkSchedule_Selected").Value                 ' Moved named cell to Configuration Tab
End Sub
Sub Update_Work_Schedule_Selection()
    Call Get_Selected_Work_Schedule
    'Hide/unhide Labor sheets based on selected Work Schedule
    Sheets(Labor_Flex980_ShName).Visible = (WorkSchedule = WorkSchedule_Flex980)
    Sheets(Labor_Flex980_2weeks_ShName).Visible = (WorkSchedule = WorkSchedule_Flex980_2weeks)
    Sheets(Labor_Flex410_ShName).Visible = (WorkSchedule = WorkSchedule_Flex410)    ' Added TSHelper 3.20
End Sub
Sub Labor_Sort_Memo_Flex980()
    Sheets(Labor_Flex980_ShName).Select
    ActiveSheet.Unprotect
    Range("B" & FirstLaborRow_Flex980 & ":O" & LastLaborRow_Flex980).Select
    Selection.Sort Key1:=Range("B" & FirstLaborRow_Flex980), _
        Order1:=xlAscending, Key2:=Range("C" & FirstLaborRow_Flex980), _
        Order2:=xlAscending, Key3:=Range("E" & FirstLaborRow_Flex980), _
        Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Call ProtectSheet(ActiveSheet.Name)
    Range("B6").Select
End Sub
Sub Labor_Sort_Memo_Flex980_2weeks()
    Sheets(Labor_Flex980_2weeks_ShName).Select
    ActiveSheet.Unprotect
    Range("B" & FirstLaborRow_Flex980_2weeks & ":AG" & LastLaborRow_Flex980_2weeks).Select
    Selection.Sort Key1:=Range("B" & FirstLaborRow_Flex980_2weeks), _
        Order1:=xlAscending, Key2:=Range("C" & FirstLaborRow_Flex980_2weeks), _
        Order2:=xlAscending, Key3:=Range("E" & FirstLaborRow_Flex980_2weeks), _
        Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Call ProtectSheet(ActiveSheet.Name)
    Range("B9").Select
End Sub
Sub Labor_Sort_Memo_Flex410() ' Added TSHelper 3.21
    Sheets(Labor_Flex410_ShName).Select
    ActiveSheet.Unprotect
    Range("B" & FirstLaborRow_Flex980 & ":O" & LastLaborRow_Flex410).Select
    Selection.Sort Key1:=Range("B" & FirstLaborRow_Flex980), _
        Order1:=xlAscending, Key2:=Range("C" & FirstLaborRow_Flex410), _
        Order2:=xlAscending, Key3:=Range("E" & FirstLaborRow_Flex410), _
        Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Call ProtectSheet(ActiveSheet.Name)
    Range("B6").Select
End Sub
Sub Labor_Sort_Workpackage_Flex980()
    Sheets(Labor_Flex980_ShName).Select
    ActiveSheet.Unprotect
    Range("B" & FirstLaborRow_Flex980 & ":O" & LastLaborRow_Flex980).Select
    Selection.Sort Key1:=Range("C" & FirstLaborRow_Flex980), _
        Order1:=xlAscending, Key2:=Range("E" & FirstLaborRow_Flex980), _
        Order2:=xlAscending, Key3:=Range("B" & FirstLaborRow_Flex980), _
        Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Call ProtectSheet(ActiveSheet.Name)
    Range("C6").Select
End Sub
Sub Labor_Sort_Workpackage_Flex980_2weeks()
    Sheets(Labor_Flex980_2weeks_ShName).Select
    ActiveSheet.Unprotect
    Range("B" & FirstLaborRow_Flex980_2weeks & ":AG" & LastLaborRow_Flex980_2weeks).Select
    Selection.Sort Key1:=Range("C" & FirstLaborRow_Flex980_2weeks), _
        Order1:=xlAscending, Key2:=Range("E" & FirstLaborRow_Flex980_2weeks), _
        Order2:=xlAscending, Key3:=Range("B" & FirstLaborRow_Flex980_2weeks), _
        Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Call ProtectSheet(ActiveSheet.Name)
    Range("C9").Select
End Sub
Sub Labor_Sort_Workpackage_Flex410() ' Added TSHelper 3.21
    Sheets(Labor_Flex410_ShName).Select
    ActiveSheet.Unprotect
    Range("B" & FirstLaborRow_Flex980 & ":O" & LastLaborRow_Flex410).Select
    Selection.Sort Key1:=Range("C" & FirstLaborRow_Flex980), _
        Order1:=xlAscending, Key2:=Range("E" & FirstLaborRow_Flex410), _
        Order2:=xlAscending, Key3:=Range("B" & FirstLaborRow_Flex410), _
        Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Call ProtectSheet(ActiveSheet.Name)
    Range("C6").Select
End Sub
Sub ClearLaborHours()
Dim result
    Call Update_Work_Schedule_Selection
    If WorkSchedule = WorkSchedule_Flex980 Then
        Call ClearLaborHours_Flex980
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        Call ShiftLaborHours_Flex980_2weeks
    Else
        Sheets(Instructions_ShName).Select
        Range("WorkSchedule_Selected").Select
        result = MsgBox("Unknown Work Schedule!", vbExclamation)
    End If
End Sub
Sub ClearLaborHours_Flex980()
Dim theDate
    Sheets(Labor_Flex980_ShName).Select
    'clear labor hours
    Range("G" & FirstLaborRow_Flex980 & ":N" & LastLaborRow_Flex980).ClearContents
    'set default days off
    Range("G8:N8").ClearContents
    Range("H8") = "X" 'Sat
    Range("I8") = "X" 'Sun
    'update Off Friday
    If UCase(Range("N7").Value) = "OFF" Then
        Range("N7") = ""
        Range("G8") = "X" 'first Fri
    Else
        Range("N7") = "Off"
        Range("N8") = "X" 'second Fri
    End If
    'update Week Ending date
    theDate = Range("K2").Value
    Range("K2").Value = theDate + 7
    'select top left cell in scrolling region (reset scroll to top left)
    Range("G10").Select
End Sub
Sub ClearLaborHours_Flex980_2weeks_Next_Week()
    Sheets(Labor_Flex980_2weeks_ShName).Select
    'clear labor hours
    Range("Y" & FirstLaborRow_Flex980_2weeks & ":AF" & LastLaborRow_Flex980_2weeks).ClearContents
    'set default days off
    Range("Y8:AF8").ClearContents
    Range("Z8") = "X" 'Sat
    Range("AA8") = "X" 'Sun
End Sub
Sub ClearLaborHours_Flex980_2weeks_Last_Week()
    Sheets(Labor_Flex980_2weeks_ShName).Select
    'clear labor hours
    Range("G" & FirstLaborRow_Flex980_2weeks & ":H" & LastLaborRow_Flex980_2weeks).ClearContents
    'set default days off
    Range("G8:H8").ClearContents
End Sub
Sub ClearLaborHours_Flex980_2weeks()
    Sheets(Labor_Flex980_2weeks_ShName).Select
    Call Clear_SAP_From_Labor_Flex980_2weeks
    'clear labor hours
    Range("M" & FirstLaborRow_Flex980_2weeks & ":T" & LastLaborRow_Flex980_2weeks).ClearContents
    'set default days off
    Range("M8:T8").ClearContents
    Range("N8") = "X" 'Sat
    Range("O8") = "X" 'Sun
    Call ClearLaborHours_Flex980_2weeks_Last_Week
    Call ClearLaborHours_Flex980_2weeks_Next_Week
    'update Days Off for Off Fridays
    If UCase(Range("T7").Value) = "OFF" Then
        Range("T8") = "X" 'second Fri this week
        Range("Y8") = "X" 'first Fri next week
    Else
        Range("M8") = "X" 'first Fri this week
        Range("AF8") = "X" 'second Fri next week
    End If
    'select top left cell in scrolling region (reset scroll to top left)
    Range("G10").Select
    'select first data entry cell for this week
    Range("M10").Select
End Sub
Sub ShiftLaborHours_Flex980_2weeks()
Dim theDate
    Sheets(Labor_Flex980_2weeks_ShName).Select
    'copy data (just Prev Friday to Thursday and Friday) from current week to last week
    '- Days off
    Range("G8") = Range("S8").Value 'Thu
    Range("H8") = Range("T8").Value 'Fri
    '- Hours
    Range("U" & FirstLaborRow_Flex980_2weeks & ":U" & LastLaborRow_Flex980_2weeks).Copy 'Total hours
    Range("G" & FirstLaborRow_Flex980_2weeks & ":G" & LastLaborRow_Flex980_2weeks).PasteSpecial Paste:=xlPasteValues
    Range("T" & FirstLaborRow_Flex980_2weeks & ":T" & LastLaborRow_Flex980_2weeks).Copy 'Friday
    Range("H" & FirstLaborRow_Flex980_2weeks & ":H" & LastLaborRow_Flex980_2weeks).PasteSpecial Paste:=xlPasteValues
    Range("G" & FirstLaborRow_Flex980_2weeks & ":G" & LastLaborRow_Flex980_2weeks).PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationSubtract 'Total - Friday
    'copy data from next week to current week
    '- Days off
    Range("Y8:AF8").Copy
    Range("M8:T8").PasteSpecial Paste:=xlPasteValues
    '- Hours
    Range("Y" & FirstLaborRow_Flex980_2weeks & ":AF" & LastLaborRow_Flex980_2weeks).Copy
    Range("M" & FirstLaborRow_Flex980_2weeks & ":T" & LastLaborRow_Flex980_2weeks).PasteSpecial Paste:=xlPasteValues
    'clear labor hours from week 2
    Call ClearLaborHours_Flex980_2weeks_Next_Week
    'update Off Friday
    If UCase(Range("T7").Value) = "OFF" Then
        Range("T7") = ""
        Range("AF8") = "X" 'second Fri next week
    Else
        Range("T7") = "Off"
        Range("Y8") = "X" 'first Fri next week
    End If
    'copy date from week 2 to week 1
    theDate = Range("AC2").Value
    Range("Q2") = theDate
    'select top left cell in scrolling region (reset scroll to top left)
    Range("G10").Select
    'select first data entry cell for this week
    Range("M10").Select
End Sub
Sub ClearLaborAdjustment()
    Sheets(Simple_Labor_Adjust_ShName).Select
    Range("C7:C16").ClearContents
    Range("H7:I16").ClearContents
    Range("C5") = 480
    Range("C5").Select
End Sub
Sub Clear_SAP_From_Labor_Flex980_2weeks()
Dim LaborRow As Long
    For LaborRow = FirstLaborRow_Flex980_2weeks To LastLaborRow_Flex980_2weeks
        If Sheets(Labor_Flex980_2weeks_ShName).Range("V" & LaborRow).Value = "S" Then
            Sheets(Labor_Flex980_2weeks_ShName).Range("B" & LaborRow & ":C" & LaborRow).ClearContents
            Sheets(Labor_Flex980_2weeks_ShName).Range("E" & LaborRow & ":H" & LaborRow).ClearContents
            Sheets(Labor_Flex980_2weeks_ShName).Range("M" & LaborRow & ":T" & LaborRow).ClearContents
            Sheets(Labor_Flex980_2weeks_ShName).Range("Y" & LaborRow & ":AF" & LaborRow).ClearContents
            Sheets(Labor_Flex980_2weeks_ShName).Range("V" & LaborRow).ClearContents
        End If
    Next
End Sub
Function Get_Work_Schedule_SheetName(WorkSchedule) As String
'returns empty string "" if WorkSchedule is invalid
    If WorkSchedule = WorkSchedule_Flex980 Then
        Get_Work_Schedule_SheetName = Labor_Flex980_ShName
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        Get_Work_Schedule_SheetName = Labor_Flex980_2weeks_ShName
    Else
        Get_Work_Schedule_SheetName = ""
    End If
End Function
Function Get_First_Labor_Row(WorkSchedule)
'returns -1 if WorkSchedule is invalid
    If WorkSchedule = WorkSchedule_Flex980 Then
        Get_First_Labor_Row = FirstLaborRow_Flex980
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        Get_First_Labor_Row = FirstLaborRow_Flex980_2weeks
    Else
        Get_First_Labor_Row = -1
    End If
End Function
Function Get_Last_Labor_Row(WorkSchedule)
'returns -1 if WorkSchedule is invalid
    If WorkSchedule = WorkSchedule_Flex980 Then
        Get_Last_Labor_Row = LastLaborRow_Flex980
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        Get_Last_Labor_Row = LastLaborRow_Flex980_2weeks
    Else
        Get_Last_Labor_Row = -1
    End If
End Function
Sub Import_GetValues_Old(theWorkbook, theSheet, colNumbers, theRow, lastrow, theValues)
'remember: colNumbers is an 8-element array of integer with column numbers for:
' 0:FRI, 1:SAT, 2:SUN, 3:MON, 4:TUE, 5:WED, 6:THU, 7:FRI(2nd Friday)
'If there isn't a 2nd Friday, set the column number to -1
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
'if theRow > lastRow then return empty strings in theValues
Dim i
    If theRow <= lastrow Then
        With Workbooks(theWorkbook).Sheets(theSheet)
            For i = 0 To 3
                theValues(i) = .Cells(theRow, i + 2).Value
            Next
            theValues(4) = .Cells(theRow, colNumbers(0)).Value  'FRI
            theValues(5) = .Cells(theRow, colNumbers(1)).Value 'SAT
            theValues(6) = .Cells(theRow, colNumbers(2)).Value 'SUN
            theValues(7) = .Cells(theRow, colNumbers(3)).Value  'MON
            theValues(8) = .Cells(theRow, colNumbers(4)).Value  'TUE
            theValues(9) = .Cells(theRow, colNumbers(5)).Value  'WED
            theValues(10) = .Cells(theRow, colNumbers(6)).Value 'THU
            If colNumbers(7) < 0 Then
                theValues(11) = "" 'leave 2nd Friday blank
            Else
                theValues(11) = .Cells(theRow, colNumbers(7)).Value '2nd Friday
            End If
        End With
    Else
        For i = 0 To 11
            theValues(i) = ""
        Next
    End If
End Sub
Sub Import_GetValues(theWorkbook, theSheet, colNumbers, theRow, lastrow, theValues)
'remember: colNumbers is an 8-element array of integer with column numbers for:
' 0:FRI, 1:SAT, 2:SUN, 3:MON, 4:TUE, 5:WED, 6:THU, 7:FRI(2nd Friday)
'If there isn't a 2nd Friday, set the column number to -1
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
'if theRow > lastRow then return empty strings in theValues
Dim i
    If theRow <= lastrow Then
        With Workbooks(theWorkbook).Sheets(theSheet)
            theValues(0) = .Cells(theRow, 2).Value
            theValues(1) = .Cells(theRow, 3).Value
            theValues(2) = .Cells(theRow, 5).Value
            theValues(3) = .Cells(theRow, 6).Value
            theValues(4) = .Cells(theRow, colNumbers(0)).Value  'FRI
            theValues(5) = .Cells(theRow, colNumbers(1)).Value  'SAT
            theValues(6) = .Cells(theRow, colNumbers(2)).Value  'SUN
            theValues(7) = .Cells(theRow, colNumbers(3)).Value  'MON
            theValues(8) = .Cells(theRow, colNumbers(4)).Value  'TUE
            theValues(9) = .Cells(theRow, colNumbers(5)).Value  'WED
            theValues(10) = .Cells(theRow, colNumbers(6)).Value 'THU
            If colNumbers(7) < 0 Then
                theValues(11) = "" 'leave 2nd Friday blank
            Else
                theValues(11) = .Cells(theRow, colNumbers(7)).Value '2nd Friday
            End If
        End With
    Else
        For i = 0 To 11
            theValues(i) = ""
        Next
    End If
End Sub
Sub Copy_GetValues(WorkSchedule, theRow, lastrow, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
'if theRow > lastRow then return empty strings in theValues
Dim theSheet
Dim i
    If theRow <= lastrow Then
        theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
        If WorkSchedule = WorkSchedule_Flex980 Then
            For i = 0 To 1
                theValues(i) = Sheets(theSheet).Cells(theRow, i + 2).Value
            Next
            For i = 2 To 11
                theValues(i) = Sheets(theSheet).Cells(theRow, i + 3).Value
            Next
        ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
            For i = 0 To 1
                theValues(i) = Sheets(theSheet).Cells(theRow, i + 2).Value
            Next
            For i = 2 To 3
                theValues(i) = Sheets(theSheet).Cells(theRow, i + 3).Value
            Next
            For i = 4 To 11
                theValues(i) = Sheets(theSheet).Cells(theRow, i + 9).Value
            Next
        End If
    Else
        For i = 0 To 11
            theValues(i) = ""
        Next
    End If
End Sub
Sub Copy_GetDaysOff(WorkSchedule, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
' We're only using 4 through 11 here
Dim theSheet
Dim i
    theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
    If WorkSchedule = WorkSchedule_Flex980 Then
        For i = 4 To 11
            theValues(i) = Sheets(theSheet).Cells(8, i + 3).Value
        Next
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        For i = 4 To 11
            theValues(i) = Sheets(theSheet).Cells(8, i + 9).Value
        Next
    End If
End Sub
Sub Cell_SetValue(theSheet, theRow, theCol, theValue)
    If theValue = "" Then
        Sheets(theSheet).Cells(theRow, theCol).ClearContents
    Else
        Sheets(theSheet).Cells(theRow, theCol).Value = theValue
    End If
End Sub
Sub Import_SetValues_Next_Week(WorkSchedule, theRow, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
Dim theSheet
Dim i, startCol
    theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
    If WorkSchedule = WorkSchedule_Flex980 Then
        'no entries for next week - nothing to do
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        'just fill in the hours for each day (4 through 11)
        For i = 4 To 11
            Call Cell_SetValue(theSheet, theRow, i + 21, theValues(i))
        Next
    End If
End Sub
Sub Import_SetValues_Last_Week(WorkSchedule, theRow, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
Dim theSheet
Dim diffValue
    theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
    If WorkSchedule = WorkSchedule_Flex980 Then
        'no entries for last week - nothing to do
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        'just fill in the hours for (the total for the week (5) - Friday (4)) and Friday (4)
        diffValue = Val(theValues(5)) - Val(theValues(4))
        If diffValue = 0 Then
            diffValue = ""
        End If
        Call Cell_SetValue(theSheet, theRow, 7, diffValue)
        Call Cell_SetValue(theSheet, theRow, 8, theValues(4))
    End If
End Sub
Sub Copy_SetValues(WorkSchedule, theRow, ByVal setAll, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
'if setAll is false, don't set the first two elements (0 and 1)
Dim theSheet
Dim i, startCol
Dim sumValue
    theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
    If setAll Then
        startCol = 0
    Else
        startCol = 2
    End If
    If WorkSchedule = WorkSchedule_Flex980 Then
        For i = startCol To 1
            Call Cell_SetValue(theSheet, theRow, i + 2, theValues(i))
        Next
        For i = 2 To 11
            Call Cell_SetValue(theSheet, theRow, i + 3, theValues(i))
        Next
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        For i = startCol To 1
            Call Cell_SetValue(theSheet, theRow, i + 2, theValues(i))
        Next
        For i = 2 To 3
            Call Cell_SetValue(theSheet, theRow, i + 3, theValues(i))
        Next
        For i = 4 To 11
            Call Cell_SetValue(theSheet, theRow, i + 9, theValues(i))
        Next
    End If
End Sub
Sub Copy_SetDaysOff(WorkSchedule, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
' We're only using 4 through 11 here
Dim theSheet
Dim i
    theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
    If WorkSchedule = WorkSchedule_Flex980 Then
        For i = 4 To 11
            Call Cell_SetValue(theSheet, 8, i + 3, theValues(i))
        Next
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        For i = 4 To 11
            Call Cell_SetValue(theSheet, 8, i + 9, theValues(i))
        Next
    End If
End Sub
Sub Import_SetDaysOff_Next_Week(WorkSchedule, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
' We're only using 4 through 11 here
Dim theSheet
Dim i
    theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
    If WorkSchedule = WorkSchedule_Flex980 Then
        'no entries for next week - nothing to do
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        For i = 4 To 11
            Call Cell_SetValue(theSheet, 8, i + 21, theValues(i))
        Next
    End If
End Sub
Sub Import_SetDaysOff_Last_Week(WorkSchedule, theValues)
'remember: theValues is a 12-element array of variant
' 0:Memo, 1:CHRG # OR CODE/Charge Object, 2:DEPT CHGD/Ext, 3:JOB CODE/Shift
' 4:FRI, 5:SAT, 6:SUN, 7:MON, 8:TUE, 9:WED, 10:THU, 11:FRI
' We're only using 4 through 11 here
Dim theSheet
Dim i
    theSheet = Get_Work_Schedule_SheetName(WorkSchedule)
    If WorkSchedule = WorkSchedule_Flex980 Then
        'no entries for next week - nothing to do
    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
        'just fill in last Friday
        Call Cell_SetValue(theSheet, 8, 7, "")  'Fri to Thu - set to blank
        Call Cell_SetValue(theSheet, 8, 8, theValues(4)) 'Fri
    End If
End Sub
Sub SetColNumbers(iFri, iSat, iSun, iMon, iTue, iWed, iThu, iFri2, colNumbers)
    colNumbers(0) = iFri
    colNumbers(1) = iSat
    colNumbers(2) = iSun
    colNumbers(3) = iMon
    colNumbers(4) = iTue
    colNumbers(5) = iWed
    colNumbers(6) = iThu
    colNumbers(7) = iFri2
End Sub
Sub Import_From_Other_Workbook()
Dim fileToOpen
Dim theError
Dim supportedFile As Boolean
Dim Continue_Import As Boolean
Dim fromWorkSchedule
Dim fromSheetName
Dim colNumbers(0 To 7) As Integer
Dim colNumbersNextWeek(0 To 7) As Integer
Dim colNumbersLastWeek(0 To 7) As Integer
Dim toFirstLaborRow, toLastLaborRow
Dim fromFirstLaborRow, fromLastLaborRow
Dim fromRow, toRow
Dim fromSuperSTAR As Boolean
Dim theValues(0 To 11) As Variant
Dim weekEndingDate
Dim importNextWeek, importLastWeek As Boolean
Dim result
Dim tempValue
Dim i
    Call Update_Work_Schedule_Selection
    'Import configuration and labor from a previous version of SuperSTAR
    DataBookName = "" 'No data workbook loaded yet
    ThisBookName = ActiveWorkbook.Name 'Remember this workbook name
    ThisSheetName = Workbooks(ThisBookName).ActiveSheet.Name 'Remember this worksheet name
    ' Present a dialog to select a file, load it and copy information over to this workbook
    fileToOpen = Application.GetOpenFilename(WorkbookFileFilter, 1, "Choose the file to import")
    If fileToOpen <> False Then
        'Error trap in case error occurs during open
        On Error Resume Next
        Workbooks.Open Filename:=fileToOpen, ReadOnly:=True 'Load the workbook file
        theError = Err.Number 'check for error: 1004 occurs if file is open already in Excel with unsaved changes
        'now turn off error trapping
        On Error GoTo 0
        If theError = 0 Then
            DataBookName = ActiveWorkbook.Name 'Keep track of new workbook name
            Workbooks(ThisBookName).Activate 'Select this workbook so we can use reference shortcuts
            Worksheets(ThisSheetName).Activate 'And make sure this same worksheet is selected (should be!)
            supportedFile = False
            Continue_Import = False
            importNextWeek = False
            importLastWeek = False
            weekEndingDate = "" 'set to blank as default
            fromSuperSTAR = False
            'check file contents to determine what import method to use
            If (Workbooks(DataBookName).Sheets.Count >= 7) Then
                If (Workbooks(DataBookName).Sheets(1).Name = "Instructions") And _
                    (Workbooks(DataBookName).Sheets(2).Name = "Labor_Flex40") And _
                    (Workbooks(DataBookName).Sheets(3).Name = "Labor_Flex980") And _
                    (Workbooks(DataBookName).Sheets(4).Name = "Labor_Flex980_2weeks") And _
                    (Workbooks(DataBookName).Sheets(5).Name = "Simple Labor Adjustment") Then
                    'file is a newer (universal) SuperSTAR file - 1.0a1 or newer
                    fromWorkSchedule = Workbooks(DataBookName).Sheets(1).Range("WorkSchedule_Selected").Value
                    fromSheetName = Get_Work_Schedule_SheetName(fromWorkSchedule)
                    If fromSheetName <> "" Then
                        With Workbooks(DataBookName).Sheets(fromSheetName)
                            If (.Range("B6").Value = "Memo") And _
                                (.Range("C6").Value = "CHRG # OR CODE") And _
                                (.Range("D6").Value = "DEPT CHGD") And _
                                (.Range("E6").Value = "JOB CODE") And _
                                (.Range("B7").Value = "Holiday") And _
                                (.Range("B8").Value = "Day off") Then
                                'valid Labor sheet (Flex 40 or Flex 9/80) version 1.0 so far
                                fromFirstLaborRow = 9
                                fromLastLaborRow = .Cells(.Rows.Count, 3).End(xlUp).Row
                                If (.Range("F6").Value = "FRI") And _
                                    (.Range("G6").Value = "SAT") And _
                                    (.Range("H6").Value = "SUN") And _
                                    (.Range("I6").Value = "MON") And _
                                    (.Range("J6").Value = "TUE") And _
                                    (.Range("K6").Value = "WED") And _
                                    (.Range("L6").Value = "THU") And _
                                    (.Range("M6").Value = "FRI") Then
                                    'This is the Flex 9/80 sheet
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = True
                                    MsgBox "*** Flex 9/80 sheet version 1.0 - import here! ***"
                                    Call SetColNumbers(6, 7, 8, 9, 10, 11, 12, 13, colNumbers)
                                ElseIf (.Range("G6").Value = "MON") And _
                                    (.Range("H6").Value = "TUE") And _
                                    (.Range("I6").Value = "WED") And _
                                    (.Range("J6").Value = "THU") And _
                                    (.Range("K6").Value = "FRI") And _
                                    (.Range("L6").Value = "SAT") And _
                                    (.Range("M6").Value = "SUN") Then
                                    'This is the Flex 40 sheet
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = True
                                    MsgBox "*** Flex 40 sheet version 1.0 - import here! ***"
                                    Call SetColNumbers(11, 12, 13, 7, 8, 9, 10, -1, colNumbers)
                                End If
                            ElseIf (.Range("B6").Value = "Memo") And _
                                (.Range("C6").Value = "CHRG # OR CODE") And _
                                (.Range("E6").Value = "DEPT CHGD") And _
                                (.Range("F6").Value = "JOB CODE") And _
                                (.Range("B7").Value = "Holiday") And _
                                (.Range("B8").Value = "Day off") Then
                                'valid Labor sheet (Flex 40 or Flex 9/80) version 1.0a6+ so far
                                fromFirstLaborRow = 9
                                fromLastLaborRow = .Cells(.Rows.Count, 3).End(xlUp).Row
                                If (.Range("G6").Value = "FRI") And _
                                    (.Range("H6").Value = "SAT") And _
                                    (.Range("I6").Value = "SUN") And _
                                    (.Range("J6").Value = "MON") And _
                                    (.Range("K6").Value = "TUE") And _
                                    (.Range("L6").Value = "WED") And _
                                    (.Range("M6").Value = "THU") And _
                                    (.Range("N6").Value = "FRI") Then
                                    'This is the Flex 9/80 sheet
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = True
                                    MsgBox "*** Flex 9/80 sheet version 1.0a6+ - import here! ***"
                                    Call SetColNumbers(7, 8, 9, 10, 11, 12, 13, 14, colNumbers)
                                ElseIf (.Range("H6").Value = "MON") And _
                                    (.Range("I6").Value = "TUE") And _
                                    (.Range("J6").Value = "WED") And _
                                    (.Range("K6").Value = "THU") And _
                                    (.Range("L6").Value = "FRI") And _
                                    (.Range("M6").Value = "SAT") And _
                                    (.Range("N6").Value = "SUN") Then
                                    'This is the Flex 40 sheet
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = True
                                    MsgBox "*** Flex 40 sheet version 1.0a6+ - import here! ***"
                                    Call SetColNumbers(12, 13, 14, 8, 9, 10, 11, -1, colNumbers)
                                End If
                            ElseIf (.Range("B9").Value = "Memo") And _
                                (.Range("C9").Value = "CHRG # OR CODE") And _
                                (.Range("D9").Value = "DEPT CHGD") And _
                                (.Range("E9").Value = "JOB CODE") And _
                                (.Range("B10").Value = "Holiday") And _
                                (.Range("B11").Value = "Day off") Then
                                'valid Labor sheet (Flex 9/80 2 weeks) version 1.0 so far
                                fromFirstLaborRow = 12
                                fromLastLaborRow = .Cells(.Rows.Count, 3).End(xlUp).Row
                                If (.Range("F9").Value = "FRI") And _
                                    (.Range("G9").Value = "SAT") And _
                                    (.Range("H9").Value = "SUN") And _
                                    (.Range("I9").Value = "MON") And _
                                    (.Range("J9").Value = "TUE") And _
                                    (.Range("K9").Value = "WED") And _
                                    (.Range("L9").Value = "THU") And _
                                    (.Range("M9").Value = "FRI") Then
                                    'This is version 1.0a1, 1.0a2, or 1.0a3
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = True
                                    Call SetColNumbers(6, 7, 8, 9, 10, 11, 12, 13, colNumbers)
                                    importNextWeek = True
                                    Call SetColNumbers(18, 19, 20, 21, 22, 23, 24, 25, colNumbersNextWeek)
                                    weekEndingDate = DateValue(.Range("I2").Value & "/" & _
                                                                .Range("K2").Value & "/" & _
                                                                .Range("M2").Value)
                                ElseIf (.Range("K9").Value = "FRI") And _
                                    (.Range("L9").Value = "SAT") And _
                                    (.Range("M9").Value = "SUN") And _
                                    (.Range("N9").Value = "MON") And _
                                    (.Range("O9").Value = "TUE") And _
                                    (.Range("P9").Value = "WED") And _
                                    (.Range("Q9").Value = "THU") And _
                                    (.Range("R9").Value = "FRI") Then
                                    'This is version - 1.0a4 or 1.0a5
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = True
                                    Call SetColNumbers(11, 12, 13, 14, 15, 16, 17, 18, colNumbers)
                                    importNextWeek = True
                                    Call SetColNumbers(23, 24, 25, 26, 27, 28, 29, 30, colNumbersNextWeek)
                                    importLastWeek = True
                                    Call SetColNumbers(6, 7, 7, 7, 7, 7, 7, 7, colNumbersLastWeek)
                                    weekEndingDate = .Range("O2").Value
                                End If
                            ElseIf (.Range("B9").Value = "Memo") And _
                                (.Range("C9").Value = "CHRG # OR CODE") And _
                                (.Range("E9").Value = "DEPT CHGD") And _
                                (.Range("F9").Value = "JOB CODE") And _
                                (.Range("B10").Value = "Holiday") And _
                                (.Range("B11").Value = "Day off") Then
                                'valid Labor sheet (Flex 9/80 2 weeks) version 1.0a6+ so far
                                fromFirstLaborRow = 12
                                fromLastLaborRow = .Cells(.Rows.Count, 3).End(xlUp).Row
                                If (.Range("M9").Value = "FRI") And _
                                    (.Range("N9").Value = "SAT") And _
                                    (.Range("O9").Value = "SUN") And _
                                    (.Range("P9").Value = "MON") And _
                                    (.Range("Q9").Value = "TUE") And _
                                    (.Range("R9").Value = "WED") And _
                                    (.Range("S9").Value = "THU") And _
                                    (.Range("T9").Value = "FRI") Then
                                    'This is version 1.0a6+
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = True
                                    Call SetColNumbers(13, 14, 15, 16, 17, 18, 19, 20, colNumbers)
                                    importNextWeek = True
                                    Call SetColNumbers(25, 26, 27, 28, 29, 30, 31, 32, colNumbersNextWeek)
                                    importLastWeek = True
                                    Call SetColNumbers(8, 9, 9, 9, 9, 9, 9, 9, colNumbersLastWeek)
                                    weekEndingDate = .Range("Q2").Value
                                End If
                            End If
                        End With
                    Else
                        Workbooks(DataBookName).Activate
                        Sheets("Instructions").Select
                        Range("WorkSchedule_Selected").Select
                        result = MsgBox("Unknown Work Schedule in file" & Chr(13) & _
                            """" & DataBookName & """!", vbExclamation)
                        supportedFile = True 'bypass "not a supported import format" error dialog
                    End If
                End If
            ElseIf (Workbooks(DataBookName).Sheets.Count = 6) Then
                If (Workbooks(DataBookName).Sheets(1).Name = "Instructions") And _
                    (Workbooks(DataBookName).Sheets(2).Name = "Labor_Flex980") And _
                    (Workbooks(DataBookName).Sheets(3).Name = "Labor_Flex980_2weeks") And _
                    (Workbooks(DataBookName).Sheets(4).Name = "Simple Labor Adjustment") Then
                    'file is an UpTEMPO file - 1.0 series
                    fromWorkSchedule = Workbooks(DataBookName).Sheets(1).Range("WorkSchedule_Selected").Value
                    fromSheetName = Get_Work_Schedule_SheetName(fromWorkSchedule)
                    If fromSheetName <> "" Then
                        With Workbooks(DataBookName).Sheets(fromSheetName)
                            If (.Range("B9").Value = "Memo") And _
                                (.Range("C9").Value = "Charge Object") And _
                                (.Range("D9").Value = "Object Length") And _
                                (.Range("E9").Value = "Ext") And _
                                (.Range("F9").Value = "Shift") Then
                                'valid Labor sheet (Flex 9/80 or Flex 9/80 2 weeks) so far
                                fromFirstLaborRow = 10
                                fromLastLaborRow = .Cells(.Rows.Count, 3).End(xlUp).Row
                                If (.Range("G9").Value = "FRI") And _
                                    (.Range("H9").Value = "SAT") And _
                                    (.Range("I9").Value = "SUN") And _
                                    (.Range("J9").Value = "MON") And _
                                    (.Range("K9").Value = "TUE") And _
                                    (.Range("L9").Value = "WED") And _
                                    (.Range("M9").Value = "THU") And _
                                    (.Range("N9").Value = "FRI") Then
                                    'This is the Flex 9/80 sheet
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = False
                                    Call SetColNumbers(7, 8, 9, 10, 11, 12, 13, 14, colNumbers)
                                    weekEndingDate = .Range("K2").Value
                                ElseIf (.Range("M9").Value = "FRI") And _
                                    (.Range("N9").Value = "SAT") And _
                                    (.Range("O9").Value = "SUN") And _
                                    (.Range("P9").Value = "MON") And _
                                    (.Range("Q9").Value = "TUE") And _
                                    (.Range("R9").Value = "WED") And _
                                    (.Range("S9").Value = "THU") And _
                                    (.Range("T9").Value = "FRI") Then
                                    'This is the Flex 9/80 2 weeks sheet
                                    supportedFile = True
                                    Continue_Import = True
                                    fromSuperSTAR = False
                                    Call SetColNumbers(13, 14, 15, 16, 17, 18, 19, 20, colNumbers)
                                    importNextWeek = True
                                    Call SetColNumbers(25, 26, 27, 28, 29, 30, 31, 32, colNumbersNextWeek)
                                    importLastWeek = True
                                    Call SetColNumbers(8, 9, 9, 9, 9, 9, 9, 9, colNumbersLastWeek)
                                    weekEndingDate = .Range("Q2").Value
                                End If
                            End If
                        End With
                    Else
                        Workbooks(DataBookName).Activate
                        Sheets("Instructions").Select
                        Range("WorkSchedule_Selected").Select
                        result = MsgBox("Unknown Work Schedule in file" & Chr(13) & _
                            """" & DataBookName & """!", vbExclamation)
                        supportedFile = True 'bypass "not a supported import format" error dialog
                    End If
                End If
            ElseIf (Workbooks(DataBookName).Sheets.Count >= 3) Then
                If (Workbooks(DataBookName).Sheets(1).Name = "Instructions") And _
                    (Workbooks(DataBookName).Sheets(2).Name = "Labor") And _
                    (Workbooks(DataBookName).Sheets(3).Name = "Simple Labor Adjustment") Then
                    'file is an older SuperSTAR file - 0.8b0 through 0.9b2
                    fromSheetName = "Labor"
                    With Workbooks(DataBookName).Sheets(2)
                        If (.Range("B6").Value = "Memo") And _
                            (.Range("C6").Value = "CHRG # OR CODE") And _
                            (.Range("D6").Value = "DEPT CHGD") And _
                            (.Range("E6").Value = "JOB CODE") And _
                            (.Range("B7").Value = "Holiday") And _
                            (.Range("B8").Value = "Day off") Then
                            'valid Labor sheet so far
                            fromFirstLaborRow = 9
                            fromLastLaborRow = .Cells(.Rows.Count, 3).End(xlUp).Row
                            If (.Range("F6").Value = "FRI") And _
                                (.Range("G6").Value = "SAT") And _
                                (.Range("H6").Value = "SUN") And _
                                (.Range("I6").Value = "MON") And _
                                (.Range("J6").Value = "TUE") And _
                                (.Range("K6").Value = "WED") And _
                                (.Range("L6").Value = "THU") And _
                                (.Range("M6").Value = "FRI") Then
                                'This is the Flex 9/80 version (0.9b1 or 0.9b2)
                                supportedFile = True
                                Continue_Import = True
                                fromSuperSTAR = True
                                Call SetColNumbers(6, 7, 8, 9, 10, 11, 12, 13, colNumbers)
                            ElseIf (.Range("G6").Value = "MON") And _
                                (.Range("H6").Value = "TUE") And _
                                (.Range("I6").Value = "WED") And _
                                (.Range("J6").Value = "THU") And _
                                (.Range("K6").Value = "FRI") And _
                                (.Range("L6").Value = "SAT") And _
                                (.Range("M6").Value = "SUN") Then
                                'This is the Flex 40 version (0.8b0 or 0.8b1 or 0.8b2)
                                supportedFile = True
                                Continue_Import = True
                                fromSuperSTAR = True
                                Call SetColNumbers(11, 12, 13, 7, 8, 9, 10, -1, colNumbers)
                            End If
                        End If
                    End With
                End If
            End If
            If Continue_Import Then
                toFirstLaborRow = Get_First_Labor_Row(WorkSchedule)
                toLastLaborRow = Get_Last_Labor_Row(WorkSchedule)
                If toFirstLaborRow < 0 Then
                    Sheets(Instructions_ShName).Select
                    Range("WorkSchedule_Selected").Select
                    result = MsgBox("Unknown Work Schedule!", vbExclamation)
                Else
                    Set_Calculation (False) 'turn off automatic calculation to speed up import
                    If fromSuperSTAR Then
                        'start with labor entries two rows up
                        fromRow = fromFirstLaborRow - 2
                    Else
                        'start with labor entries on this row
                        fromRow = fromFirstLaborRow
                        'but first, transfer Day Off values
                        Call Import_GetValues(DataBookName, fromSheetName, colNumbers, fromRow - 2, fromLastLaborRow, theValues)
                        Call Copy_SetDaysOff(WorkSchedule, theValues)
                    End If
                    toRow = toFirstLaborRow
                    Do While toRow <= toLastLaborRow
                        Call Import_GetValues(DataBookName, fromSheetName, colNumbers, fromRow, fromLastLaborRow, theValues)
                        If fromSuperSTAR Then
                            'blank out DEPT CHARGED and JOB CODE
                            theValues(2) = ""
                            theValues(3) = ""
                            'check for Ext (3 digits at end of 15-character charge number)
                            If Len(theValues(1)) > 12 Then
                                If UCase(Left(theValues(1), 3)) = "P00" Then
                                    'SAP charge number - leave alone
                                Else
                                    'copy 3-character extension to Ext
                                    theValues(2) = Mid(theValues(1), 13, 3)
                                    'set the Charge Object to 12 characters
                                    theValues(1) = Left(theValues(1), 12)
                                End If
                            End If
                        End If
                        If theValues(1) = "DO" Then 'special case - map "Day Off" (CHRG CODE "DO") hours to days off indicators
                            'change days with hours to "X" strings
                            For i = 4 To 11
                                If Not (theValues(i) = "") Then
                                    theValues(i) = "X"
                                End If
                            Next
                            Call Copy_SetDaysOff(WorkSchedule, theValues)
                            fromRow = fromRow + 1
                        Else
                            Call Copy_SetValues(WorkSchedule, toRow, toRow >= toFirstLaborRow, theValues)
                            fromRow = fromRow + 1
                            toRow = toRow + 1
                        End If
                    Loop
                    If WorkSchedule = WorkSchedule_Flex980 Then
                        If weekEndingDate <> "" Then
                            'set current week ending date
                            Sheets(Get_Work_Schedule_SheetName(WorkSchedule)).Range("K2") = weekEndingDate
                        End If
                    ElseIf WorkSchedule = WorkSchedule_Flex980_2weeks Then
                        If weekEndingDate <> "" Then
                            'set current week ending date
                            Sheets(Get_Work_Schedule_SheetName(WorkSchedule)).Range("Q2") = weekEndingDate
                        End If
                        If importNextWeek Then
                            'import data for next week
                            If fromSuperSTAR Then
                                fromRow = fromFirstLaborRow - 2
                            Else
                                fromRow = fromFirstLaborRow
                                Call Import_GetValues(DataBookName, fromSheetName, colNumbersNextWeek, fromRow - 2, fromLastLaborRow, theValues)
                                Call Import_SetDaysOff_Next_Week(WorkSchedule, theValues)
                            End If
                            toRow = toFirstLaborRow
                            Do While toRow <= toLastLaborRow
                                Call Import_GetValues(DataBookName, fromSheetName, colNumbersNextWeek, fromRow, fromLastLaborRow, theValues)
                                If theValues(1) = "DO" Then 'special case - map "Day Off" (CHRG CODE "DO") hours to days off indicators
                                    'change days with hours to "X" strings
                                    For i = 4 To 11
                                        If Not (theValues(i) = "") Then
                                            theValues(i) = "X"
                                        End If
                                    Next
                                    Call Import_SetDaysOff_Next_Week(WorkSchedule, theValues)
                                    fromRow = fromRow + 1
                                Else
                                    Call Import_SetValues_Next_Week(WorkSchedule, toRow, theValues)
                                    fromRow = fromRow + 1
                                    toRow = toRow + 1
                                End If
                            Loop
                        Else
                            'clear data for next week
                            Call ClearLaborHours_Flex980_2weeks_Next_Week
                        End If
                        If importLastWeek Then
                            'import data for last week
                            If fromSuperSTAR Then
                                fromRow = fromFirstLaborRow - 2
                            Else
                                fromRow = fromFirstLaborRow
                                Call Import_GetValues(DataBookName, fromSheetName, colNumbersLastWeek, fromRow - 2, fromLastLaborRow, theValues)
                                Call Import_SetDaysOff_Last_Week(WorkSchedule, theValues)
                            End If
                            toRow = toFirstLaborRow
                            Do While toRow <= toLastLaborRow
                                Call Import_GetValues(DataBookName, fromSheetName, colNumbersLastWeek, fromRow, fromLastLaborRow, theValues)
                                If theValues(1) = "DO" Then 'special case - map "Day Off" (CHRG CODE "DO") hours to days off indicators
                                    'change days with hours to "X" strings
                                    For i = 4 To 11
                                        If Not (theValues(i) = "") Then
                                            theValues(i) = "X"
                                        End If
                                    Next
                                    Call Import_SetDaysOff_Last_Week(WorkSchedule, theValues)
                                    fromRow = fromRow + 1
                                Else
                                    Call Import_SetValues_Last_Week(WorkSchedule, toRow, theValues)
                                    fromRow = fromRow + 1
                                    toRow = toRow + 1
                                End If
                            Loop
                        Else
                            'clear data for last week
                            Call ClearLaborHours_Flex980_2weeks_Last_Week
                        End If
                    End If
                    Set_Calculation (True) 'turn automatic calculation back on
                    'Now copy settings over
                    tempValue = Workbooks(DataBookName).Sheets("Instructions").Range("AllLabor_X").Value
                    Sheets(Instructions_ShName).Range("AllLabor_X").Value = tempValue
                End If
            End If
            Import_Close_Workbook
            ' Let user know we're done
            Workbooks(ThisBookName).Activate 'Select original workbook
            Worksheets(ThisSheetName).Activate 'And original worksheet
            If Not supportedFile Then
                result = MsgBox("File """ & DataBookName & """ is not a supported import format." & Chr(13) & Chr(13) & _
                    "Please try again using an UpTEMPO file or SuperSTAR file (version 0.8b0 or newer)", vbExclamation)
            End If
            If Continue_Import Then
                Sheets(Get_Work_Schedule_SheetName(WorkSchedule)).Select
                result = MsgBox("Successfully imported data from file" & Chr(13) & _
                    """" & DataBookName & """.", vbInformation)
            End If
        End If
    End If
End Sub
Private Sub Import_Close_Workbook()
'
' Closes the data workbook
'
    If DataBookName <> "" Then
        Workbooks(DataBookName).Worksheets(1).Range("A1").Copy 'Copy a single cell to empty large buffer
        Workbooks(DataBookName).Close savechanges:=False
    End If
End Sub
Sub Copy_From_Other_Work_Schedule()
Dim fromWorkSchedule
Dim toFirstLaborRow, toLastLaborRow
Dim fromFirstLaborRow, fromLastLaborRow
Dim fromRow, toRow
Dim theValues(12) As Variant
Dim result
    Call Update_Work_Schedule_Selection
    'Copy labor from a different work schedule
    fromWorkSchedule = Range("WorkSchedule_CopyFrom").Value
    If fromWorkSchedule = WorkSchedule Then
        Range("WorkSchedule_CopyFrom").Select
        result = MsgBox("The 'Copy from' Work Schedule is the same as the currently selected Work Schedule.  No copying is necessary!", vbExclamation)
    Else
        toFirstLaborRow = Get_First_Labor_Row(WorkSchedule)
        toLastLaborRow = Get_Last_Labor_Row(WorkSchedule)
        If toFirstLaborRow < 0 Then
            Sheets(Instructions_ShName).Select
            Range("WorkSchedule_Selected").Select
            result = MsgBox("Unknown Work Schedule!", vbExclamation)
            Exit Sub
        End If
        fromFirstLaborRow = Get_First_Labor_Row(fromWorkSchedule)
        fromLastLaborRow = Get_Last_Labor_Row(fromWorkSchedule)
        If fromFirstLaborRow < 0 Then
            Sheets(Instructions_ShName).Select
            Range("WorkSchedule_CopyFrom").Select
            result = MsgBox("Unknown Work Schedule!", vbExclamation)
            Exit Sub
        End If
        fromRow = fromFirstLaborRow
        toRow = toFirstLaborRow
        Set_Calculation (False) 'turn off automatic calculation to speed up copy
        Call Copy_GetDaysOff(fromWorkSchedule, theValues)
        Call Copy_SetDaysOff(WorkSchedule, theValues)
        Do While toRow <= toLastLaborRow
            Call Copy_GetValues(fromWorkSchedule, fromRow, fromLastLaborRow, theValues)
            Call Copy_SetValues(WorkSchedule, toRow, toRow >= toFirstLaborRow, theValues)
            fromRow = fromRow + 1
            toRow = toRow + 1
        Loop
        Set_Calculation (True) 'turn automatic calculation back on
        result = MsgBox("Copy complete!", vbInformation)
    End If
End Sub
Private Sub UnprotectSheets()
    'Make sure this workbook is the active workbook
    Call Workbook_Activate
    'Unhide sheets that are normally hidden
    Sheets(Labor_Flex980_ShName).Visible = True
    Sheets(Labor_Flex980_2weeks_ShName).Visible = True
    Sheets(Dropdown_Entries_ShName).Visible = True
    'Unprotect sheets for updates
    Sheets(Instructions_ShName).Unprotect
    Sheets(Labor_Flex980_ShName).Unprotect
    Sheets(Labor_Flex980_2weeks_ShName).Unprotect
    Sheets(Simple_Labor_Adjust_ShName).Unprotect
    Sheets(Dropdown_Entries_ShName).Unprotect
End Sub
Sub ProtectSheet(sheetName)
    Sheets(sheetName).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
End Sub
Sub ProtectSheets()
    'Make sure this workbook is the active workbook
    Call Workbook_Activate
    'Make sure sheets are protected
    Sheets(Instructions_ShName).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Call ProtectSheet(Labor_Flex980_ShName)
    Call ProtectSheet(Labor_Flex980_2weeks_ShName)
    Call ProtectSheet(Simple_Labor_Adjust_ShName)
    Sheets(Dropdown_Entries_ShName).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    'Also hide sheets that should be hidden
    Sheets(Dropdown_Entries_ShName).Visible = False
    Call Update_Work_Schedule_Selection
End Sub
Sub CleanForDistribution()
' Clears data to prepare workbook for distribution to others
    'Make sure this workbook is the active workbook
    Call Workbook_Activate
    'Make sure sheets are protected
    Call ProtectSheets
    'Unhide sheets to clear them
    Sheets(Labor_Flex980_ShName).Visible = True
    Sheets(Labor_Flex980_2weeks_ShName).Visible = True
    Sheets(Dropdown_Entries_ShName).Visible = True
    'Clear contents of Labor Adjustment sheet
    Call ClearLaborAdjustment
    'Use Flex980_2weeks as master date and schedule
    'copy Last Week Ending date from Flex980_2weeks sheet to "current" date on Flex980 sheet
    Sheets(Labor_Flex980_ShName).Range("K2").Value = Sheets(Labor_Flex980_2weeks_ShName).Range("H2").Value
    'copy Last Week's Off Friday status from Flex980_2weeks sheet to Flex980 sheet
    Sheets(Labor_Flex980_ShName).Range("N7").Value = Sheets(Labor_Flex980_2weeks_ShName).Range("H7").Value
    'Clear contents of Flex980 sheet (also increments date and toggles Off Friday)
    Call ClearLaborHours_Flex980
    'Clear contents of Flex980_2weeks sheet
    Call ClearLaborHours_Flex980_2weeks
    'Make sure sheets are protected (also re-hides sheets that should be hidden)
    Call ProtectSheets
    'Clear the workpackage information
    Sheets(Labor_Flex980_ShName).Range("B" & FirstLaborRow_Flex980 & ":C" & LastLaborRow_Flex980).ClearContents
    Sheets(Labor_Flex980_ShName).Range("E" & FirstLaborRow_Flex980 & ":F" & LastLaborRow_Flex980).ClearContents
    Sheets(Labor_Flex980_2weeks_ShName).Range("B" & FirstLaborRow_Flex980_2weeks & ":C" & LastLaborRow_Flex980_2weeks).ClearContents
    Sheets(Labor_Flex980_2weeks_ShName).Range("E" & FirstLaborRow_Flex980_2weeks & ":F" & LastLaborRow_Flex980_2weeks).ClearContents
    'Reset the goals to 40 hours
    Sheets(Labor_Flex980_ShName).Range("F7") = 40
    Sheets(Labor_Flex980_2weeks_ShName).Range("F7") = 40
    'Set default Configuration values on Instructions sheet
    Sheets(Instructions_ShName).Select
    Range("TEMPO_URL").Value = Default_URL_TEMPO
    Range("TEMPO_ShellHome_Suffix").Value = Default_Suffix_Shell_Home
    Range("TEMPO_TimeEntry_Suffix").Value = Default_Suffix_Time_Entry
    Range("TEMPO_LoggedOff_URL").Value = Default_URL_LoggedOff
    Range("AllLabor_X").Value = ""
    Range("MacroWarning_X").Value = "X"
    Range("CompletedDialog_X").Value = "X"
    Range("Timeout_Delay").Value = DefaultTimeOut
    Range("Single_Delay").Value = DefaultSingleDelay
    Range("Double_Delay").Value = DefaultDoubleDelay
    'Set scroll and selected cell on Instructions sheet
    Sheets(Instructions_ShName).Select
    Range("A3").Select 'scroll the instructions sheet to the top (the section below the frozen top 2 rows)
    Range("A2").Select 'then put the selection on row 2
End Sub

'Code Module SHA-512
'''e42ceeece9ed14e67479cde9e3482b5e35a72a05126f733da556bd1e78d6e1bf380f9c470fad7917673f05c918c973d42f8104efba444e90069d53fbb990df75