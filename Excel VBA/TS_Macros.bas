' *******************************************
' Timesheet Macros
'
' Version:
' 3.00_b2 - Brian Whipp
'
' ********************************************

' Developer Notes:
'1. UpTEMPO Current Version - 1.0a3 (2016-09-13)
'2. Only changes made to William's Code:
'   InternetExplorerObjects Module
'   Was:         WEdate = Sheets(CallingSheet).Range("K2").Value
'   Changed to:  WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
'
'   Was:         WEdate = Sheets(CallingSheet).Range("Q2").Value
'   Changed to:  WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
'3. Sheet Labor_Flex980
'   Changed N5 to =BH10
'4. Sheet Labor_Flex980_2weeks
'   Changed T5 to =BH10

Public TS_MaxRows As Long
Public Lastrow As Long
Public BottomRow As Long
Public Const TS_MaxDefaultRows = 2000
Public Resize As Boolean


Private Const DevMode = "Dev_Mode"

Sub TS_UpdateMaxRows()

' Find current maximum number used

Resize = True

'Turn off automatic calculation

Application.Calculation = xlCalculationManual

Call TS_CalcMaxRows
Call TS_UpdateMaxNames
Call TS_ClearTimesheet

'If TS_MaxRows <= Lastrow Then
'    Resize = False
'    Application.Calculation = xlCalculationAutomatic
'    Exit Sub
'Else
'    'Set up the names ranges
'    Call TS_UpdateMaxNames
'    Call TS_ClearTimesheet
'End If

Application.Calculation = xlCalculationAutomatic

End Sub

Sub TS_CalcMaxRows()

Dim sht As Worksheet

Set sht = ThisWorkbook.Worksheets("Timesheet")

TS_MaxRows = Sheets("Configuration").Range("E20").Value

'Ctrl + Shift + End
  Lastrow = sht.Cells(sht.Rows.Count, "C").End(xlUp).Row + 1
  BottomRow = sht.Cells(sht.Rows.Count, "O").End(xlUp).Row + 1

If TS_MaxRows < Lastrow Then
    Sheets("Configuration").Range("E20").Value = Lastrow + 2
    TS_MaxRows = Lastrow + 2
End If

End Sub

Sub TS_UpdateMaxNames()

' This will add it to the Name Manager, so should run this once or on any change
' Part of set up and the eventual change capability
' For change, make sure we don't allow to shrink beyond existing lines

' Formulas that factor off of Timesheet length
' Configuration:    Q = Timesheet - A, H, I
' Summary:  C4 = Timesheet - H, I
'           C5 = Timesheet - H, I
'           C6 = Timesheet - H, I
'           C7 = Timesheet - H, I
' Timesheet:    T = O, AM, AA
'               U = AA, T
'               V = AB, U
'               AD = P
'               AE = P
'               AF = P
'               AG = P
'               AH = P
' Labor_Flex980:    BI = Timesheet - R
'                   BJ = Timesheet - R
'                   G-N = Timesheet - T, R, AI, S
' WP Lookup:    BH-DG = Timesheet - T, S, R
' Clean for Distribution:   A, B, C, D, E, G, M
' Conditional Formatting:  E, J, L


' So, A, B, C, D, E, G, H, I, J, L, M, O, P, R, S, T, U, AA, AB, AI, AM

 ActiveWorkbook.Names.Add _
      Name:="TS_Amax", _
      RefersTo:="=Timesheet!$A$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Bmax", _
      RefersTo:="=Timesheet!$B$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Cmax", _
      RefersTo:="=Timesheet!$C$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Dmax", _
      RefersTo:="=Timesheet!$D$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Emax", _
      RefersTo:="=Timesheet!$E$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Gmax", _
      RefersTo:="=Timesheet!$G$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Hmax", _
      RefersTo:="=Timesheet!$H$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Imax", _
      RefersTo:="=Timesheet!$I$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Jmax", _
      RefersTo:="=Timesheet!$J$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Lmax", _
      RefersTo:="=Timesheet!$L$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Mmax", _
      RefersTo:="=Timesheet!$M$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Omax", _
      RefersTo:="=Timesheet!$O$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Pmax", _
      RefersTo:="=Timesheet!$P$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Rmax", _
      RefersTo:="=Timesheet!$R$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Smax", _
      RefersTo:="=Timesheet!$S$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Tmax", _
      RefersTo:="=Timesheet!$T$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_Umax", _
      RefersTo:="=Timesheet!$U$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_AAmax", _
      RefersTo:="=Timesheet!$AA$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_ABmax", _
      RefersTo:="=Timesheet!$AB$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_AImax", _
      RefersTo:="=Timesheet!$AI$" & TS_MaxRows

 ActiveWorkbook.Names.Add _
      Name:="TS_AMmax", _
      RefersTo:="=Timesheet!$AM$" & TS_MaxRows


End Sub

Sub TS_DeveloperMode()
' Toggle Developer Mode
' Hide/Unhide Sheets that are normally hidden
' Unprotect all Sheets
' Update States

If Range(DevMode) = "Off" Then
    Call TS_DevMode_On
  Else
    Call TS_DevMode_Off
End If

End Sub

Sub TS_DevMode_On()
' Turn on Developer Mode

Call TS_UnprotectSheets

Sheets("Configuration").Select
Range("A20").Value = "On"

Call TS_UnhideSheets
Sheets("Configuration").Activate

End Sub
Sub TS_DevMode_Off()
' Turn on Developer Mode

Sheets("Configuration").Select
Range("A20").Value = "Off"

Call TS_ProtectSheets
Call TS_HideSheets

End Sub

Sub TS_ProtectSheets()
'Make sure sheets are protected
    
    Sheets("LM Command Media").Protect
    Sheets("Directions").Protect
    Sheets("Instructions").Protect
    Sheets("Configuration").Protect
    Sheets("Summary").Protect
    Sheets("Timesheet").Protect
    Sheets("Labor_Flex980").Protect
    Sheets("Labor_Flex980_2weeks").Protect
    Sheets("WP #'s").Protect
    Sheets("WP Lookup").Protect
    Sheets("Change History").Protect
    Sheets("Dropdown_Entries").Protect
    Sheets("TSMasterFormulas").Protect

' Don't touch the Macro Warning sheet
' Otherwise, William's macros will break
'   Sheets("Macro Warning").Protect

End Sub

Sub TS_UnprotectSheets()
'Make sure sheets are unprotected
    
    Sheets("LM Command Media").Unprotect
    Sheets("Directions").Unprotect
    Sheets("Instructions").Unprotect
    Sheets("Configuration").Unprotect
    Sheets("Summary").Unprotect
    Sheets("Timesheet").Unprotect
    Sheets("Labor_Flex980").Unprotect
    Sheets("Labor_Flex980_2weeks").Unprotect
    Sheets("WP #'s").Unprotect
    Sheets("WP Lookup").Unprotect
    Sheets("Change History").Unprotect
    Sheets("Dropdown_Entries").Unprotect
    Sheets("TSMasterFormulas").Unprotect

' Don't touch the Macro Warning sheet
' Otherwise, William's macros will break
'   Sheets("Macro Warning").Unprotect

End Sub

Sub TS_HideSheets()
' Hide Sheets for normal use
    
    Sheets("Dropdown_Entries").Visible = xlSheetHidden
    Sheets("Macro Warning").Visible = xlSheetHidden
    Sheets("ExecutionTimes").Visible = xlSheetHidden
    Sheets("TSMasterFormulas").Visible = xlSheetHidden
    Call Update_Work_Schedule_Selection

End Sub

Sub TS_UnhideSheets()
' Unhide Sheets for developer use
    
    Sheets("Labor_Flex980").Visible = True
    Sheets("Labor_Flex980_2weeks").Visible = True
    Sheets("Dropdown_Entries").Visible = True
    Sheets("Macro Warning").Visible = True
    Sheets("ExecutionTimes").Visible = True
    Sheets("TSMasterFormulas").Visible = True

End Sub


Sub TS_CleanForDistribution()
' Clears data to prepare workbook for distribution to others
' Modified for Timesheet

    Resize = False

    'Make sure sheets are unprotected
    Call TS_UnprotectSheets
    
    Application.Calculation = xlCalculationManual

    Call TS_ClearLM_Command_Media
    Call TS_ClearDirections
    Call TS_ClearInstructions
    Call TS_ClearConfiguration
    
    ' Must redo because Developer Mode resets the protections
    Call TS_UnprotectSheets
    Call TS_UnhideSheets
    
    Call TS_ClearSummary
    Call TS_ClearTimesheet
    Call TS_ClearLabor_Flex980
    Call TS_ClearLabor_Flex980_2weeks
    Call TS_ClearWPs
    
    'Make sure sheets are protected
    Call TS_ProtectSheets
    Call TS_HideSheets
    
    ' Reset to front page
    Call TS_ClearLM_Command_Media
    
Application.Calculation = xlCalculationAutomatic

End Sub
Sub TS_ClearLM_Command_Media()

    Sheets("LM Command Media").Select

' Return to Corner
    Range("A1").Select

End Sub

Sub TS_ClearDirections()

    Sheets("Directions").Select

' Return to Corner
    Range("A1").Select

End Sub

Sub TS_ClearInstructions()

    Sheets("Instructions").Select

' Quick Start Section
    Range("E10").Value = "Flex 9/80"                ' Flex 9/80 Default
    Range("E19").Value = "Flex 9/80"                ' Flex 9/80 Default

' Configuration
    Range("F22").Value = "https://tempofdb.external.lmco.com/fiori"    ' TEOMPO URL
    Range("G24").Select                     ' Enter all lines of labor
    Selection.ClearContents
    Range("G26").Value = "X"                ' Enable Macro Warning

' Web Interface Tuning
    Range("C61").Value = "15"               ' Timeout
    Range("C63").Value = "1"                ' Delay
    Range("C65").Value = "2"                ' Double Delay
    
' Return to Corner
    Range("A1").Select

End Sub

Sub TS_ClearConfiguration()
' Clear/Reset the few values

    Sheets("Configuration").Select
    
' Reset Values
    Range("A2").Value = "6"                 ' End of the Week Day
    Range("C2").Select                      ' Worksheet Year
    Selection.ClearContents
    Range("E2").Select                      ' Vacation Accrued
    Selection.ClearContents
    Range("G2").Select                      ' Vacation Hours at Start of Year
    Selection.ClearContents
    Range("I2").Select                      ' Floating Holidays for Year
    Selection.ClearContents
    Range("K2").Select                      ' Holiday Hours for Year
    Selection.ClearContents
    Range("A20").Value = "On"              ' Take it out of Deevloper Mode
    Call TS_DeveloperMode
    Sheets("Configuration").Range("E20").Value = TS_MaxDefaultRows   ' Set Max Rows

' Return to Worksheet Year
    
    Range("C2").Select

End Sub

Sub TS_ClearSummary()
' Clear/Reset the one value

    Sheets("Summary").Select
    
' Reset Value
    Range("C25").Select                      ' Charge Number Lookup
    Selection.ClearContents

End Sub

Sub TS_ClearTimesheet()

' This can be used as a way to expand automatically, just need to make sure set up row is correct offsets

' Clear all but first (header) row
Dim tempRange As String
Dim startRow As Double
Dim startRow2 As Double    ' The row before

If Not Resize Then
    Call TS_CalcMaxRows
    Call TS_UpdateMaxNames
End If

Sheets("Timesheet").Select

If Resize Then
    ' Start from LastRow to BottomRow
    tempRange = "A" & Lastrow & ":AR" & BottomRow
    Range(tempRange).Select
Else
    ' Clear from row 2 to BottomRow
    Range("A2:AR" & BottomRow).Select
End If

Selection.ClearContents

' Copy Master Formulas from TSMasterFormulas sheet
' Then Fill down

' Odd behavior, tyring to do full line will results in
' an error - an array error with format painter.
' Need to break it up into two sections

Sheets("TSMasterFormulas").Select

Range("A2:N2").Select
Selection.Copy

Sheets("Timesheet").Select

' Set up "first" row
' Will need to do all cells in the row
    
With Sheets("Timesheet").Range("A" & Lastrow)
    .PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With

Sheets("TSMasterFormulas").Select

Range("O2:AR2").Select
Selection.Copy

Sheets("Timesheet").Select

' Set up "first" row
' Will need to do all cells in the row
    
With Sheets("Timesheet").Range("O" & Lastrow)
    .PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With


If TS_MaxRows < 2 Then
    TS_MaxRows = 2
End If

If TS_MaxRows < Lastrow Then
    TS_MaxRows = Lastrow
End If

If Resize Then
    ' Start from LastRow to BottomRow
    tempRange = "A" & Lastrow & ":AR" & TS_MaxRows
    Worksheets("Timesheet").Range(tempRange).FillDown
Else
    Worksheets("Timesheet").Range("A2:AR" & TS_MaxDefaultRows).FillDown
End If

    Range("A2").Select                      ' Place cursor back at home

End Sub

Sub TS_ClearLabor_Flex980()
' Reset the Week Ending Date, everything else is Protected

    Sheets("Labor_Flex980").Select
    
' Reset Values
    Range("K2").Value = "10/2/2016"          ' Payroll Week Ending Date
    Range("F7").Value = "40"                 ' Hours goal for week
    Range("G8").Select                       ' Fri On
    Selection.ClearContents
    Range("H8").Value = "X"                  ' Sat Off
    Range("I8").Value = "X"                  ' Sun Off
    Range("J8").Select                       ' Mon On
    Selection.ClearContents
    Range("K8").Select                       ' Tue On
    Selection.ClearContents
    Range("L8").Select                       ' Wed On
    Selection.ClearContents
    Range("M8").Select                       ' Thur On
    Selection.ClearContents
    Range("N8").Select                       ' Fri On
    Selection.ClearContents
    
    Range("K2").Select                      ' Place cursor back at home
    
End Sub

Sub TS_ClearLabor_Flex980_2weeks()
' Reset the Week Ending Date, everything else is Protected

    Sheets("Labor_Flex980_2weeks").Select
    
' Reset Values
    Range("Q2").Value = "10/2/2016"          ' Payroll Week Ending Date
    Range("F7").Value = "40"                 ' Hours goal for week
    Range("M8").Select                       ' Fri On
    Selection.ClearContents
    Range("N8").Value = "X"                  ' Sat Off
    Range("O8").Value = "X"                  ' Sun Off
    Range("P8").Select                       ' Mon On
    Selection.ClearContents
    Range("Q8").Select                       ' Tue On
    Selection.ClearContents
    Range("R8").Select                       ' Wed On
    Selection.ClearContents
    Range("S8").Select                       ' Thur On
    Selection.ClearContents
    Range("T8").Select                       ' Fri On
    Selection.ClearContents
    
    Range("Q2").Select                      ' Place cursor back at home

End Sub


Sub TS_ClearWPs()
' Reset the Work Package Listings

    Sheets("WP #'s").Select
    
' Set Intial Values
    Range("A2").Value = "-"             ' WP #
    Range("B2").Value = "NOTE"          ' Shortcut
    Range("A3").Value = "-"             ' WP #
    Range("B3").Value = "Break"          ' Shortcut
    Range("A4").Value = "-"             ' WP #
    Range("B4").Value = "Breakfast"          ' Shortcut
    Range("A5").Value = "-"             ' WP #
    Range("B5").Value = "Lunch"          ' Shortcut
    Range("A6").Value = "-"             ' WP #
    Range("B6").Value = "Dinner"          ' Shortcut
    Range("A7").Value = "V"             ' WP #
    Range("B7").Value = "Vacation"          ' Shortcut
    Range("A8").Value = "PI"             ' WP #
    Range("B8").Value = "Sick"          ' Shortcut
    Range("A9").Value = "H"             ' WP #
    Range("B9").Value = "Holiday"          ' Shortcut
    Range("A10").Value = "HF"             ' WP #
    Range("B10").Value = "Floating Holiday"          ' Shortcut
    
    Range("C2:G150").Select
    Selection.ClearContents
    
    Range("A11:A150").Value = "_blank_"
    
'    For i = 11 To 150
'        Sheets("WP #'s").Cells(i, 1).Value = "_blank_"
'    Next i

    Range("A2").Select                  ' Put cursor back at home

End Sub

