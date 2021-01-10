'Subroutines for use by all modules
Option Explicit
'
Public HatchCols As Long
Public HatchRows As Long
'
'
Sub Set_Row_Col_Hatch_Ranges()
Dim sheetName
    sheetName = ActiveSheet.Name
    If sheetName = Instructions_ShName Then
        HatchCols = 11
        HatchRows = Find_Last_Row(3)
    ElseIf sheetName = Labor_Flex980_ShName Then
        HatchCols = 16
        HatchRows = LastLaborRow_Flex980 + 1
    ElseIf sheetName = Labor_Flex980_2weeks_ShName Then
        HatchCols = 34
        HatchRows = LastLaborRow_Flex980_2weeks + 1
    ElseIf sheetName = Simple_Labor_Adjust_ShName Then
        HatchCols = 15
        HatchRows = 19
    ElseIf sheetName = Dropdown_Entries_ShName Then
        HatchCols = 3
        HatchRows = Find_Last_Row(3)
        If HatchRows < 3 Then HatchRows = 3
    Else 'unknown sheet - don't hatch!
        HatchCols = 0
        HatchRows = 0
    End If
End Sub
Sub Hatch_Locked_Cells()
Dim iRow As Long
Dim iCol As Long
    Call Set_Row_Col_Hatch_Ranges
    For iRow = 1 To HatchRows
        For iCol = 1 To HatchCols
            If Cells(iRow, iCol).Locked Then
                Cells(iRow, iCol).Interior.Pattern = xlLightDown
            End If
        Next iCol
    Next iRow
End Sub
Sub Hatch_Unlocked_Cells()
Dim iRow As Long
Dim iCol As Long
    Call Set_Row_Col_Hatch_Ranges
    For iRow = 1 To HatchRows
        For iCol = 1 To HatchCols
            If Not Cells(iRow, iCol).Locked Then
                Cells(iRow, iCol).Interior.Pattern = xlLightUp
            End If
        Next iCol
    Next iRow
End Sub
Sub UnHatch_All_Cells()
Dim iRow As Long
Dim iCol As Long
    Call Set_Row_Col_Hatch_Ranges
    If (ActiveSheet.Name = Instructions_ShName) Or _
       (ActiveSheet.Name = Dropdown_Entries_ShName) Then
        For iRow = 1 To HatchRows
            For iCol = 1 To HatchCols
                If Cells(iRow, iCol).Interior.Pattern <> xlNone Then
                    Cells(iRow, iCol).Interior.Pattern = xlNone
                End If
            Next iCol
        Next iRow
    Else
        For iRow = 1 To HatchRows
            For iCol = 1 To HatchCols
                If Cells(iRow, iCol).Interior.Pattern <> xlSolid Then
                    Cells(iRow, iCol).Interior.Pattern = xlSolid
                End If
            Next iCol
        Next iRow
    End If
End Sub
Function Find_Last_Row(theColumn)
    Find_Last_Row = Cells(Rows.Count, theColumn).End(xlUp).Row
End Function
Function Find_Last_Column(theRow)
    Find_Last_Column = Cells(theRow, Columns.Count).End(xlToLeft).Column
End Function
Sub Excel_Activate()
' Activates Excel - brings Excel window (and message box, if open) to front
Dim appWindowTitle
' This no longer works in Excel 2013 (also started failing in Excel 2010 on 9/21/2015):
'    AppActivate ("Microsoft Excel")
' This often works in Excel 2013, but sometimes fails:
'    AppActivate ("Excel")
' So we use this new method:
' (note that without an object qualifier, Application represents the entire Microsoft Excel application)
    appWindowTitle = Application.Caption    ' Get the current Excel window title
    AppActivate (appWindowTitle)            ' Use AppActivate with full window title
' Temporary debug:
'    MsgBox appWindowTitle
End Sub
Sub Set_Calculation(turnOn)
    If turnOn Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub


