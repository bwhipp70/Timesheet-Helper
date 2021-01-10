' *******************************************
' Timesheet Macros
'
' Version:
' 3.00 - Brian Whipp, based on UpTEMPO 1.0a3 (2016-09-13)
' 3.01 - Brian Whipp, included updates from UpTEMPO 1.0a4 (2016-09-23)
' 3.02 - Brian Whipp
'
' ********************************************

' Developer Notes:
'1. UpTEMPO Current Version - 1.0a4 (2016-09-23)
'2. Only changes made to William's Code:
'   InternetExplorerObjects Module
'   Was:         WEdate = Sheets(CallingSheet).Range("K2").Value
'   Changed to:  WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
'
'   Was:         WEdate = Sheets(CallingSheet).Range("Q2").Value
'   Changed to:  WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
'
'******************************************
'* Everything Else is addinbg PtrSafe to allow it to run on 64 bit machines
'******************************************
' Was to:
' Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal _
'  hwnd As Long) As Long
'
' Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal _
'  hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'
' Changed to:
' Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal _
'  hwnd As Long) As Long
'
' Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal _
'  hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'
'   Subroutines:
'   Was:
'   Private Declare Function BringWindowToTop Lib "user32" (ByVal _
'    hwnd As Long) As Long
'
'   Private Declare Function FindWindow Lib "user32" Alias _
'    "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName _
'    As Any) As Long
'
'   Private Declare Function GetTopWindow Lib "user32" (ByVal _
'    hwnd As Long) As Long
'
'   Private Declare Function IsIconic Lib "user32" (ByVal _
'    hwnd As Long) As Long
'
'   Private Declare Function OpenIcon Lib "user32" (ByVal _
'    hwnd As Long) As Long

'   Changed to:
'   Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal _
'    hwnd As Long) As Long
'
'   Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
'    "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName _
'    As Any) As Long
'
'   Private Declare PtrSafe Function GetTopWindow Lib "user32" (ByVal _
'    hwnd As Long) As Long
'
'   Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal _
'    hwnd As Long) As Long
'
'   Private Declare PtrSafe Function OpenIcon Lib "user32" (ByVal _
'    hwnd As Long) As Long
'
'   KeyboardState:
'   Was:
' Private Declare Function GetVersionEx Lib "kernel32" _
'    Alias "GetVersionExA" _
'    (lpVersionInformation As OSVERSIONINFO) As Long
'
' Private Declare Sub keybd_event Lib "user32" _
'    (ByVal bVk As Byte, _
'     ByVal bScan As Byte, _
'     ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'
' Private Declare Function GetKeyboardState Lib "user32" _
'    (pbKeyState As Byte) As Long
'
' Private Declare Function SetKeyboardState Lib "user32" _
'   (lppbKeyState As Byte) As Long

'   Changed to:
' Private Declare PtrSafe Function GetVersionEx Lib "kernel32" _
'    Alias "GetVersionExA" _
'    (lpVersionInformation As OSVERSIONINFO) As Long
'
' Private Declare PtrSafe Sub keybd_event Lib "user32" _
'    (ByVal bVk As Byte, _
'     ByVal bScan As Byte, _
'     ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'
' Private Declare PtrSafe Function GetKeyboardState Lib "user32" _
'    (pbKeyState As Byte) As Long
'
' Private Declare PtrSafe Function SetKeyboardState Lib "user32" _
'   (lppbKeyState As Byte) As Long

'3. Sheet Labor_Flex980
'   Changed N5 to =BH10
'4. Sheet Labor_Flex980_2weeks
'   Changed T5 to =BH10
'5. WP #'s - Column I was originally sorted, removing duplicates and blanks.  This slows things down considerably.
'   Name Manager - WP_List =OFFSET('WP #''s'!$A$2,0,0,COUNTA('WP #''s'!$A$2:$A$150))
'   Name Manager - WP_List_Unique_alpha =OFFSET('WP #''s'!$I$2, 0, 0, COUNT(IF('WP #''s'!$I$2:$I$149="", "", 1)), 1)
'   WP #'s, Column I, cell I:2 {=IFERROR(INDEX(WP_List, MATCH(0, IF(MAX(NOT(COUNTIF($I$1:I1, WP_List))*(COUNTIF(WP_List, ">"&WP_List)+1))=(COUNTIF(WP_List, ">"&WP_List)+1), 0, 1), 0)),"")}
'
'   Timing Values, COlumn I, Baseline = 1.88005
'   Timing Values, WP_List = A2:A150 = 1.71694
'   Timing Values, WP_List_Alpha_Unique = Column I = 2.0026
'   Numbers don't make much sense, performance is noticeably better?!

' Option Explicit broke the Import Capability
' Option Explicit

' Variables used to resize the Timesheet
Public TS_MaxRows As Long
Public lastRow As Long
Public BottomRow As Long
Public Const TS_MaxDefaultRows = 2000
Public Resize As Boolean

'Workbook and worksheet variables used by import routines
Private DataBookName
Private ThisBookName
Private ThisSheetName

Private Const DevMode = "Dev_Mode"


Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long


Sub TS_OpenTEMPO()

' Brings up TEMPO
'
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", "https://tempo.external.lmco.com/fiori")

End Sub


Sub TS_UpdateMaxRows()

' Entry for the Macro Button
' Determines the maximum used entries by looking for last # in Column C
' And also determines overall size of sheet by last formula line

' Find current maximum number used

Resize = True

'Turn off automatic calculation

Application.Calculation = xlCalculationManual

Call TS_CalcMaxRows         ' Determine the last used row and last formula row
Call TS_UpdateMaxNames      ' Set up the names to the new last row in the Name Manager
Call TS_ClearTimesheet      ' Adjust Timesheet

Application.Calculation = xlCalculationAutomatic

End Sub

Sub TS_CalcMaxRows()

Dim sht As Worksheet

Set sht = ThisWorkbook.Worksheets("Timesheet")

' User field, find desired length
' TS_MaxRows = Sheets("Configuration").Range("E20").Value
TS_MaxRows = Sheets("Configuration").Range("AdjustRows").Value

'Ctrl + Shift + End to find last used rows
  lastRow = sht.Cells(sht.Rows.Count, "C").End(xlUp).Row + 1        ' Last user row
  BottomRow = sht.Cells(sht.Rows.Count, "O").End(xlUp).Row + 1      ' Last Forumla row

' If the user requested length would remove entires, set to max entries + 2 to provide buffer
' Don't let them remove entries

If TS_MaxRows < lastRow Then
'    Sheets("Configuration").Range("E20").Value = lastRow + 2
    Sheets("Configuration").Range("AdjustRows").Value = lastRow + 2
    TS_MaxRows = lastRow + 2
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
'           D15 = Timesheet - V
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


' So, A, B, C, D, E, G, H, I, J, L, M, O, P, R, S, T, U, V, AA, AB, AI, AM

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
      RefersTo:="=Timesheet!$F$" & TS_MaxRows

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
      Name:="TS_Vmax", _
      RefersTo:="=Timesheet!$V$" & TS_MaxRows

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

' Set the flag for the user
Sheets("Configuration").Select
' Range("A20").Value = "On"
Range("Dev_Mode").Value = "On"

Call TS_UnhideSheets
Sheets("Configuration").Activate

End Sub
Sub TS_DevMode_Off()
' Turn on Developer Mode

' Set the flag for the user
Sheets("Configuration").Select
'Range("A20").Value = "Off"
Range("Dev_Mode").Value = "Off"

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
    
    Sheets("Instructions").Visible = xlSheetHidden
    Sheets("Dropdown_Entries").Visible = xlSheetHidden
    Sheets("Macro Warning").Visible = xlSheetHidden
    Sheets("ExecutionTimes").Visible = xlSheetHidden
    Sheets("TSMasterFormulas").Visible = xlSheetHidden
    Call Update_Work_Schedule_Selection     ' This will hide either the Labor_Flex980 or Labor_Flex980_2weeks sheet

End Sub

Sub TS_UnhideSheets()
' Unhide Sheets for developer use
    
    Sheets("Instructions").Visible = True
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
    Call TS_ClearConfiguration
    
    ' Must redo because Developer Mode resets the protections
    Call TS_UnprotectSheets
    Call TS_UnhideSheets
    
    Call TS_ClearInstructions
    
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
'    Range("E10").Value = "Flex 9/80"                ' Flex 9/80 Default
    Range("WorkSchedule_Selected").Value = "Flex 9/80"                ' Flex 9/80 Default
'    Range("E19").Value = "Flex 9/80"                ' Flex 9/80 Default
    Range("WorkSchedule_CopyFrom").Value = "Flex 9/80"                ' Flex 9/80 Default

' Configuration
'    Range("F22").Value = "https://tempofdb.external.lmco.com/fiori"    ' TEOMPO URL
    Range("TEMPO_URL").Value = "https://tempo.external.lmco.com/fiori"    ' TEOMPO URL
'    Range("G24").Select                     ' Enter all lines of labor
    Range("AllLabor_X").Select                     ' Enter all lines of labor
    Selection.ClearContents
'    Range("G26").Value = "X"                ' Enable Macro Warning
    Range("MacroWarning_X").Value = "X"                ' Enable Macro Warning
'    Range("I28").Value = "X"
    Range("CompletedDialog_X").Value = "X"

' Web Interface Tuning
'    Range("C63").Value = "15"               ' Timeout
'    Range("C65").Value = "1"                ' Delay
'    Range("C67").Value = "2"                ' Double Delay
    Range("Timeout_Delay").Value = "15"               ' Timeout
    Range("Single_Delay").Value = "1"                ' Delay
    Range("Double_Delay").Value = "2"                ' Double Delay
    Range("TEMPO_ShellHome_Suffix").Value = "#Shell-home"  ' Suffix for Shell Home Page
    Range("TEMPO_TimeEntry_Suffix").Value = "#ZTPOTIMESHEET3-record"  ' Suffix for Time Entry Page
    Range("TEMPO_LoggedOff_URL").Value = "https://tempo.external.lmco.com/sap/public/bc/icf/logoff" ' URL for TEMPO logged off
    
' Return to Corner
    Range("A1").Select

End Sub

Sub TS_ClearConfiguration()
' Clear/Reset the few values

    Sheets("Configuration").Select
    
' Reset Values
'    Range("A2").Value = "6"                 ' End of the Week Day
    Range("EndoftheWeekDay").Value = "6"                 ' End of the Week Day
'    Range("C2").Select                      ' Worksheet Year
    Range("WS_Year").Select                      ' Worksheet Year
    Selection.ClearContents
'    Range("E2").Select                      ' Vacation Accrued
    Range("VacationAccrued").Select                      ' Vacation Accrued
    Selection.ClearContents
'    Range("G2").Select                      ' Vacation Hours at Start of Year
    Range("VacationStart").Select                      ' Vacation Hours at Start of Year
    Selection.ClearContents
'    Range("I2").Select                      ' Floating Holidays for Year
    Range("FloatHolidays").Select                      ' Floating Holidays for Year
    Selection.ClearContents
'    Range("K2").Select                      ' Holiday Hours for Year
    Range("HolidayHrs").Select                      ' Holiday Hours for Year
    Selection.ClearContents
'    Range("A20").Value = "On"              ' Take it out of Deevloper Mode
    Range("Dev_Mode").Value = "On"              ' Take it out of Deevloper Mode
'    Sheets("Configuration").Range("E20").Value = TS_MaxDefaultRows   ' Set Max Rows
    Sheets("Configuration").Range("AdjustRows").Value = TS_MaxDefaultRows   ' Set Max Rows

'    Range("H19").Value = "X"                 ' Import Configuration
'    Range("H20").Value = "X"                 ' Import WP #'s
'    Range("H21").Value = "X"                 ' Import Timesheet
    Range("ImportConfig").Value = "X"                 ' Import Configuration
    Range("ImportWP").Value = "X"                 ' Import WP #'s
    Range("ImportTimesheet").Value = "X"                 ' Import Timesheet
    
    Range("WP_Dropdown").Select             ' Clear extra WP drop down list sorting
    Selection.ClearContents

    Range("M:W").EntireColumn.Hidden = True  ' Hide Columns
    
    Columns("B").ColumnWidth = 1.57 ' 16 pixels
    Columns("D").ColumnWidth = 1.57 ' 16 pixels
    Columns("F").ColumnWidth = 1.57 ' 16 pixels
    Columns("H").ColumnWidth = 1.57 ' 16 pixels
    Columns("J").ColumnWidth = 1.57 ' 16 pixels
    Columns("L").ColumnWidth = 1.57 ' 16 pixels
    
    Columns("A").ColumnWidth = 21.57 '156 pixels
    Columns("C").ColumnWidth = 10.14 '76 pixels
    Columns("E").ColumnWidth = 20#   '145 pixels
    Columns("G").ColumnWidth = 14.57 '107 pixels
    Columns("I").ColumnWidth = 14.57 '107 pixels
    Columns("K").ColumnWidth = 14.57 '107 pixels
    Columns("X").ColumnWidth = 8.43 '64 pixels
    
    Range("A:X").Font.Name = "Calibri"
    Range("A:X").Font.Size = 11

' Reset Developer Mode
    Call TS_DeveloperMode

' Return to Worksheet Year
    
    Range("C2").Select

End Sub

Sub TS_ClearSummary()
' Clear/Reset the one value

    Sheets("Summary").Select
    
    Columns("A").ColumnWidth = 1.29 '14 pixels
    Columns("B").ColumnWidth = 19.57 '142 pixels
    Columns("C").ColumnWidth = 13.29 '98 pixels
    Columns("D").ColumnWidth = 17.57 '128 pixels
    
    Range("A:D").Font.Name = "Calibri"
    Range("A:D").Font.Size = 11
   
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
Dim tsProtected As Boolean
Dim tsMasterVisible As Boolean

tsProtected = Sheets("Timesheet").ProtectContents

If tsProtected Then
    Sheets("Timesheet").Unprotect
End If

tsMasterVisible = Sheets("TSMasterFormulas").Visible

If Not tsMasterVisible Then
    Sheets("TSMasterFormulas").Visible = True
End If


If Not Resize Then
    Call TS_CalcMaxRows
    Call TS_UpdateMaxNames
End If

Sheets("Timesheet").Select

If Resize Then
    ' Start from LastRow to BottomRow
    tempRange = "A" & lastRow & ":AR" & BottomRow
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

' If Clean for Distribution, start at the very top
If Not Resize Then
    lastRow = 2
End If
    

Sheets("TSMasterFormulas").Select

Range("A2:N3").Select
Selection.Copy

Sheets("Timesheet").Select

' Set up "first" row
' Will need to do all cells in the row
    
With Sheets("Timesheet").Range("A" & lastRow)
    .PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With

Sheets("TSMasterFormulas").Select

Range("O2:AR3").Select
Selection.Copy

Sheets("Timesheet").Select

' Set up "first" row
' Will need to do all cells in the row
    
With Sheets("Timesheet").Range("O" & lastRow)
    .PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With


If TS_MaxRows < 3 Then
    TS_MaxRows = 3
End If

If TS_MaxRows < lastRow Then
    TS_MaxRows = lastRow
End If

If Resize Then
    ' Start from LastRow to BottomRow
    tempRange = "A" & lastRow & ":AR" & TS_MaxRows
    Worksheets("Timesheet").Range(tempRange).FillDown
Else
    Worksheets("Timesheet").Range("A3:AR" & TS_MaxDefaultRows).FillDown
End If

Sheets("Timesheet").Select
Range("A1048576:AR1048576").Select
Selection.Copy

If Resize Then
    With Sheets("Timesheet").Range("A" & TS_MaxRows + 1)
        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
    tempRange = "A" & TS_MaxRows + 1 & ":AR20000"
    Worksheets("Timesheet").Range(tempRange).FillDown

Else
    With Sheets("Timesheet").Range("A" & TS_MaxDefaultRows + 1)
        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
    tempRange = "A" & TS_MaxDefaultRows + 1 & ":AR20000"
    Worksheets("Timesheet").Range(tempRange).FillDown
End If

' Set up Blocker / Reminder Rows

Range("G2").Select
Selection.Copy
   
If Resize Then
   
    tempRange = "A" & TS_MaxRows + 1 & ":M" & TS_MaxRows + 3
   
    With Sheets("Timesheet").Range(tempRange)
       .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With

    Range("D" & TS_MaxRows + 2).Value = "*** STOP - Use the Adjust Rows to Add More Lines"

Else
    tempRange = "A" & TS_MaxDefaultRows + 1 & ":M" & TS_MaxDefaultRows + 3
   
    With Sheets("Timesheet").Range(tempRange)
       .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With

    Range("D" & TS_MaxDefaultRows + 2).Value = "*** STOP - Use the Adjust Rows to Add More Lines"
End If


    Range("O:AR").EntireColumn.Hidden = True  ' Hide Columns
    
    Columns("A").ColumnWidth = 12#   '89 pixels
    Columns("B").ColumnWidth = 4.57 '37 pixels
    Columns("C").ColumnWidth = 4.57 '37 pixels
    Columns("D").ColumnWidth = 49.71 '353 pixels
    Columns("E").ColumnWidth = 12.86 '95 pixels
    Columns("F").ColumnWidth = 5.86 '46 pixels
    Columns("G").ColumnWidth = 4.57 '37 pixels
    Columns("H").ColumnWidth = 12.86 '95 pixels
    Columns("I").ColumnWidth = 5.43 '43 pixels
    Columns("J").ColumnWidth = 4.57 '37 pixels
    Columns("K").ColumnWidth = 5.43 '43 pixels
    Columns("L").ColumnWidth = 7.86 '60 pixels
    Columns("M").ColumnWidth = 26.71 '192 pixels
    
    Range("A:AR").Font.Name = "Consolas"
    Range("A:AR").Font.Size = 8

'If Timesheet was originally protected, reprotect
If tsProtected Then
    Sheets("Timesheet").Protect
End If

If Not tsMasterVisible Then
    Sheets("TSMasterFormulas").Visible = xlSheetHidden
End If
    
    
    Range("A2").Select                      ' Place cursor back at home

End Sub

Sub TS_ClearLabor_Flex980()
' Reset the Week Ending Date, everything else is Protected

    Sheets("Labor_Flex980").Select
    
' Reset Values
    Range("K2").Value = "=TODAY()"           ' Payroll Week Ending Date
    Range("F7").Value = "40"                 ' Hours goal for week
    Range("G8").Formula = "=IF(G6="""",""X"","""")"  ' Fri Auto Select
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
    Range("N8").Formula = "=IF(N6="""",""X"","""")"   ' Fri Auto Select
    
    Range("K2").Select                      ' Place cursor back at home
    
End Sub

Sub TS_ClearLabor_Flex980_2weeks()
' Reset the Week Ending Date, everything else is Protected

    Sheets("Labor_Flex980_2weeks").Select
    
' Reset Values
    Range("Q2").Value = "=TODAY()"           ' Payroll Week Ending Date
    Range("F7").Value = "40"                 ' Hours goal for week
    Range("M8").Formula = "=IF(M6="""",""X"","""")"    ' Fri Auto Select
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
    Range("T8").Formula = "=IF(T6="""",""X"","""")"   ' Fri Auto Select
    
    Range("Q2").Select                      ' Place cursor back at home

End Sub


Sub TS_ClearWPs()
' Reset the Work Package Listings

    Sheets("WP #'s").Select
    
' Set Intial Values
    Range("A2").Value = "-"             ' WP #
    Range("B2").Value = "NOTE"          ' Shortcut
    Range("C2").Value = "Placeholder for a Note (No WP)"          ' Description
  
    Range("A3").Value = "-"             ' WP #
    Range("B3").Value = "Break"          ' Shortcut
    Range("C3").Value = "Placeholder for a Break (No WP)"          ' Description
    
    Range("A4").Value = "-"             ' WP #
    Range("B4").Value = "Breakfast"          ' Shortcut
    Range("C4").Value = "Placeholder for Breakfast (No WP)"          ' Description
    
    Range("A5").Value = "-"             ' WP #
    Range("B5").Value = "Lunch"          ' Shortcut
    Range("C5").Value = "Placeholder for Lunch (No WP)"          ' Description
    
    Range("A6").Value = "-"             ' WP #
    Range("B6").Value = "Dinner"          ' Shortcut
    Range("C6").Value = "Placeholder for Dinner (No WP)"          ' Description
    
    Range("A7").Value = "PA"             ' WP #
    Range("B7").Value = "Vacation"          ' Shortcut
    Range("C7").Value = "Vacation - Accrued Paid Time Off (PA)"          ' Description
    
    Range("A8").Value = "PG"             ' WP #
    Range("B8").Value = "Sick"          ' Shortcut
    Range("C8").Value = "Sick Time / Personal Business - Granted Paid Time Off (PG)"          ' Description
    
    Range("A9").Value = "PS"             ' WP #
    Range("B9").Value = "Holiday"          ' Shortcut
    Range("C9").Value = "Holiday - Fixed Paid Time Off (PS)"          ' Description
    
    Range("A10").Value = "PF"             ' WP #
    Range("B10").Value = "Floating Holiday"          ' Shortcut
    Range("C10").Value = "Floating Holiday - Floating Paid Time Off (PF)"          ' Description
    
    
' 2018 Training Charge Numbers
    
    Range("A11").Value = "SC"             ' WP #
    Range("B11").Value = ""          ' Shortcut
    Range("C11").Value = "Security Training"          ' Description
    
    Range("A12").Value = "TC"             ' WP #
    Range("B12").Value = ""          ' Shortcut
    Range("C12").Value = "Ethics & Business Compliance (Ex: BCCT)"          ' Description
    
    Range("A13").Value = "TR          REQ"             ' WP #
    Range("B13").Value = ""          ' Shortcut
    Range("C13").Value = "Corporate or RMS Required Training (Ex: Import/Export, CAM)"          ' Description
    
    Range("A14").Value = "TR          MGT"             ' WP #
    Range("B14").Value = ""          ' Shortcut
    Range("C14").Value = "Leadership Training (CLE,FSL)"          ' Description

    Range("A15").Value = "TR          FEL"             ' WP #
    Range("B15").Value = ""          ' Shortcut
    Range("C15").Value = "LM Fellows Conference"          ' Description
    
    Range("A16").Value = "TR          EPD"             ' WP #
    Range("B16").Value = ""          ' Shortcut
    Range("C16").Value = "New Development Enterprise Product Data Management"          ' Description
    
    Range("A17").Value = "TR          EPM"             ' WP #
    Range("B17").Value = ""          ' Shortcut
    Range("C17").Value = "Engineering Project Management Training"          ' Description

    Range("A18").Value = "TR          TOP"             ' WP #
    Range("B18").Value = ""          ' Shortcut
    Range("C18").Value = "Top Gun (IWSS, TLS, SAC-Helo, C4USS-TBD)"          ' Description

    Range("A19").Value = "TR          CYB"             ' WP #
    Range("B19").Value = ""          ' Shortcut
    Range("C19").Value = "Cyber Training"          ' Description

    Range("A20").Value = "TR          DEV"             ' WP #
    Range("B20").Value = ""          ' Shortcut
    Range("C20").Value = "Course Development"          ' Description

    Range("A21").Value = "TR          NEO"             ' WP #
    Range("B21").Value = ""          ' Shortcut
    Range("C21").Value = "New Employee Orientation"          ' Description

    Range("A22").Value = "TR          CON"             ' WP #
    Range("B22").Value = ""          ' Shortcut
    Range("C22").Value = "Conferences"          ' Description

    Range("A23").Value = "TR          TRN"             ' WP #
    Range("B23").Value = ""          ' Shortcut
    Range("C23").Value = "General Technical Training"          ' Description

    Range("A24").Value = "TR          DDA"             ' WP #
    Range("B24").Value = ""          ' Shortcut
    Range("C24").Value = "DDE - Agile"          ' Description

    Range("A25").Value = "TR          DDT"             ' WP #
    Range("B25").Value = ""          ' Shortcut
    Range("C25").Value = "DDE - Automation"         ' Description

    Range("A26").Value = "TR          DDM"             ' WP #
    Range("B26").Value = ""          ' Shortcut
    Range("C26").Value = "DDE - Model Based Engineering"         ' Description

    Range("A27").Value = "TR          DDE"             ' WP #
    Range("B27").Value = ""          ' Shortcut
    Range("C27").Value = "Digital Development Environment - DDE Other"         ' Description

    Range("A28").Value = "TR          ADQ"             ' WP #
    Range("B28").Value = ""          ' Shortcut
    Range("C28").Value = "Architect Development & Qualification Pgm"         ' Description

    Range("A29").Value = "TR          SED"             ' WP #
    Range("B29").Value = ""          ' Shortcut
    Range("C29").Value = "Systems Engineering Development & Qualification Pgm"         ' Description

    Range("A30").Value = "TR          ESP"             ' WP #
    Range("B30").Value = ""          ' Shortcut
    Range("C30").Value = "Embedded Systems Program"         ' Description

    Range("A25").Value = "TR          COR"             ' WP #
    Range("B25").Value = ""          ' Shortcut
    Range("C25").Value = "Training Coordination/Administration"         ' Description

    
    Range("B26:B150").Select
    Selection.ClearContents
   
    Range("C26:G150").Select
    Selection.ClearContents
    
    Range("A26:A150").Value = "_blank_"
    
    Range("I:I").EntireColumn.Hidden = True  ' Hide Columns
    
    Columns("A").ColumnWidth = 14.14 '104 pixels
    Columns("B").ColumnWidth = 14.14 '104 pixels
    Columns("C").ColumnWidth = 76.14 '538 pixels
    Columns("D").ColumnWidth = 17.57 '128 pixels
    Columns("E").ColumnWidth = 8.57  '65 pixels
    Columns("F").ColumnWidth = 10.14 '76 pixels
    Columns("G").ColumnWidth = 24.86 '179 pixels
    
    Range("A1:G1").Font.Name = "Calibri"
    Range("A1:G1").Font.Size = 13
    
    Range("A2:G150").Font.Name = "Consolas"
    Range("A2:G150").Font.Size = 8
    
    Range("A2").Select                  ' Put cursor back at home

End Sub

Sub TS_ImportData()

' Import selected data
' Stolen heavily from William Hall's code

Dim fileToOpen
Dim theError
Dim configSh As Boolean
Dim wpSh As Boolean
Dim timesheetSh As Boolean
Dim v208 As Boolean
Dim v209 As Boolean
Dim v300 As Boolean
Dim TSVer As Double
Dim sht As Worksheet
Dim SheetExists As Boolean
Dim theValues(0 To 11) As Variant
Dim i
Dim importBottomRow As Long
Dim testString As String
Dim tsProtected As Boolean
Dim v208len As Long

' Grab sheet protection state to restore it later.
' Must unprotect the sheet to make changes.
tsProtected = Sheets("Timesheet").ProtectContents

If tsProtected Then
    Sheets("Timesheet").Unprotect
End If

' Set variables to determine which version to import
v208 = False
v209 = False
v300 = False
TSVer = 0

' Dummy string to validate if workbook to import is correct
testString = ""

' Determine which sheets to import
configSh = False
wpSh = False
timesheetSh = False

Sheets("Configuration").Select

'If (Range("H19").Value <> "") Then
If (Range("ImportConfig").Value <> "") Then
    configSh = True
End If

'If (Range("H20").Value <> "") Then
If (Range("ImportWP").Value <> "") Then
    wpSh = True
End If

'If (Range("H21").Value <> "") Then
If (Range("ImportTimesheet").Value <> "") Then
    timesheetSh = True
End If

    'Import configuration and labor from a previous version of Timesheet
    
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
            DataBookName = ActiveWorkbook.Name      'Keep track of new workbook name
            Workbooks(ThisBookName).Activate        'Select this workbook so we can use reference shortcuts
            Worksheets(ThisSheetName).Activate      'And make sure this same worksheet is selected (should be!)
        End If
   
        ' Configuration and WP #'s will be version based
        ' Transfer in appropriate data
                       
        SheetExists = False
        
        ' Make sure it is a valid workbook
        ' If not, trap the error and gracefully inform the user
        
        On Error Resume Next
        Workbooks(DataBookName).Activate
        testString = Worksheets("Change History").Range("A1")
        On Error GoTo 0
        If testString = "Rev History" Then
            SheetExists = True      ' Sheet Exists, so probably a good import candidate
        Else
            SheetExists = False     ' Sheet does not exist, not a good import candidate
        End If
        
        If Not SheetExists Then
            Call TS_Import_Close_Workbook
            MsgBox "Change History sheet not found in " & DataBookName & ", please select a different file."
            'If Timesheet was originally protected, reprotect
                If tsProtected Then
                    Sheets("Timesheet").Protect
                End If
            Exit Sub
        End If
        
        
        ' Looks like a good workbook, so try to determine what the version is
        ' Layout and functionality changes were made in 2.08/2.09/3.00
        
        
        If (Workbooks(DataBookName).Sheets("Change History").Range("A49") = "Rev 2.08") Then
            v208 = True
            TSVer = 2.08
        End If
        
        If (Workbooks(DataBookName).Sheets("Change History").Range("A61") = "Rev 2.09") Then
            v209 = True
            TSVer = 2.09
        End If
        
        If (Workbooks(DataBookName).Sheets("Change History").Range("A93") = "Rev 3.00") Then
            v300 = True
            TSVer = 3
        End If
            
        If (Workbooks(DataBookName).Sheets("Configuration").Range("X1") = "Version") Then
            If (Workbooks(DataBookName).Sheets("Configuration").Range("X2") = "3.02.01") Then
                TSVer = 3.02
            Else
                TSVer = Sheets("Configuration").Range("X2").Value
            End If
        End If
           
'        MsgBox "SheetExists = " & SheetExists & "; V208/209/300 = " & v208 & ";" & v209 & ";" & v300
'        MsgBox "TSVer = " & TSVer
        
'       Needs to be V2.08 or newer to import.

        If (Not v208) And (Not v209) And (Not v300) And (Not (TSVer > 3)) Then
            Call TS_Import_Close_Workbook
'            MsgBox "Invalid Workbook File, Please Try Again"
            result = MsgBox("File """ & DataBookName & """ is not a supported import format." & Chr(13) & Chr(13) & _
                    "Please try again using a Timesheet file (version 2.08 or newer)", vbExclamation)
            'If Timesheet was originally protected, reprotect
                If tsProtected Then
                    Sheets("Timesheet").Protect
                End If
            Exit Sub
        End If
        
' Disable calculation to speed things up
' Wanted to do this after validity check so as to not accidentally leave it off

    Application.Calculation = xlCalculationManual
        
' Import the Configuration Tab values
        If configSh Then
        ' A2, C2, E2, G2, I2, K2 always copy
        ' A20, E20 only if V3.00 or greater
        
           Workbooks(DataBookName).Activate
           Sheets("Configuration").Select
           theValues(0) = Sheets("Configuration").Range("A2").Value
           theValues(1) = Sheets("Configuration").Range("C2").Value
           theValues(2) = Sheets("Configuration").Range("E2").Value
           theValues(3) = Sheets("Configuration").Range("G2").Value
           theValues(4) = Sheets("Configuration").Range("I2").Value
           theValues(5) = Sheets("Configuration").Range("K2").Value
         
           If v300 Then
             theValues(6) = Sheets("Configuration").Range("A20").Value
             theValues(7) = Sheets("Configuration").Range("E20").Value
           End If
         
' Moved some items around in 3.02
           If TSVer >= 3.02 Then
             theValues(6) = Sheets("Configuration").Range("A11").Value
             theValues(7) = Sheets("Configuration").Range("E11").Value
           End If
         
           Workbooks(ThisBookName).Activate
           Sheets("Configuration").Select
'           Sheets("Configuration").Range("A2").Value = theValues(0)
'           Sheets("Configuration").Range("C2").Value = theValues(1)
'           Sheets("Configuration").Range("E2").Value = theValues(2)
'           Sheets("Configuration").Range("G2").Value = theValues(3)
'           Sheets("Configuration").Range("I2").Value = theValues(4)
'           Sheets("Configuration").Range("K2").Value = theValues(5)
           Sheets("Configuration").Range("EndoftheWeekDay").Value = theValues(0)
           Sheets("Configuration").Range("WS_Year").Value = theValues(1)
           Sheets("Configuration").Range("VacationAccrued").Value = theValues(2)
           Sheets("Configuration").Range("VacationStart").Value = theValues(3)
           Sheets("Configuration").Range("FloatHolidays").Value = theValues(4)
           Sheets("Configuration").Range("HolidayHrs").Value = theValues(5)
         
           If v300 Then
'             Sheets("Configuration").Range("A20").Value = theValues(6)
'             Sheets("Configuration").Range("E20").Value = theValues(7)
             Sheets("Configuration").Range("Dev_Mode").Value = theValues(6)
             Sheets("Configuration").Range("AdjustRows").Value = theValues(7)
           End If
     
' Return to Worksheet Year
    
           Range("C2").Select
     
        End If
        
' Import the WP #'s Tab values

        If wpSh Then
        ' ChangeHistory A49 = 2.08 or skip
        ' A2:G150 if B = Shortcut
        ' A2:A150, B2:F150 if B = Description
          If v300 Then
             Workbooks(DataBookName).Activate
             Sheets("WP #'s").Select
             ActiveSheet.Range("A2:G150").Select
             Selection.Copy
             Workbooks(ThisBookName).Activate
             Sheets("WP #'s").Select
             ActiveSheet.Range("A2:G150").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          Else
            If (v209 And Not v300) Then
               Workbooks(DataBookName).Activate
               Sheets("WP #'s").Select
               ActiveSheet.Range("A2:A150").Select
               Selection.Copy
               Workbooks(ThisBookName).Activate
               Sheets("WP #'s").Select
               ActiveSheet.Range("A2:A150").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
               Workbooks(DataBookName).Activate
               Sheets("WP #'s").Select
               ActiveSheet.Range("B2:F150").Select
               Selection.Copy
               Workbooks(ThisBookName).Activate
               Sheets("WP #'s").Select
               ActiveSheet.Range("C2:G150").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Else
                If v208 Then
                   Workbooks(DataBookName).Activate
                   Sheets("Configuration").Select
                   
                   ' Unprotect the v208 sheet
                   ActiveSheet.Unprotect "56o$sdfH"
                   
                   ' Figure out how many lines are used in the shortcut
                   v208len = Range("N2:N20").Cells.SpecialCells(xlCellTypeConstants).Count
                   ' Copy only the used rows
                   ActiveSheet.Range("N2:N" & 2 + v208len - 1).Select
                   Selection.Copy
                   Workbooks(ThisBookName).Activate
                   Sheets("WP #'s").Select
                   ActiveSheet.Range("A2:A" & 2 + v208len - 1).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                   
                   Workbooks(DataBookName).Activate
                   Sheets("Configuration").Select
                   ActiveSheet.Range("M2:M" & 2 + v208len - 1).Select
                   Selection.Copy
                   Workbooks(ThisBookName).Activate
                   Sheets("WP #'s").Select
                   ActiveSheet.Range("B2:B" & 2 + v208len - 1).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                   ' Copy all WPs up to the maximum number of free rows
                   Workbooks(DataBookName).Activate
                   Sheets("WP #'s").Select
                   ActiveSheet.Range("A2:A" & 150 - v208len).Select
                   Selection.Copy
                   Workbooks(ThisBookName).Activate
                   Sheets("WP #'s").Select
                   ActiveSheet.Range("A" & 2 + v208len & ":A150").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
                   Workbooks(DataBookName).Activate
                   Sheets("WP #'s").Select
                   ActiveSheet.Range("B2:F" & 150 - v208len).Select
                   Selection.Copy
                   Workbooks(ThisBookName).Activate
                   Sheets("WP #'s").Select
                   ActiveSheet.Range("C" & 2 + v208len & ":G150").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                End If
            
            End If
          End If
        
        Range("A2").Select                  ' Put cursor back at home
        
        End If
        
' Import the Timesheet entries

        If timesheetSh Then
        ' Copy everything down to last row
        ' v3.05 and before, Move Columns A-E, G, L-M
        ' v3.06 and after, Move Columns A-F, L-M
        ' Determine number of total rows and used rows in source sheet
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
        ' Ctrl + Shift + End
             importBottomRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "O").End(xlUp).Row + 1
             
        ' Set up destination sheet length
             Workbooks(ThisBookName).Activate
             Sheets("Configuration").Select
'             Range("E20").Value = importBottomRow - 1
             Range("AdjustRows").Value = importBottomRow - 1
             Call TS_UpdateMaxRows
             
        ' Copy data
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("A2:E" & importBottomRow).Select
             Selection.Copy
             Workbooks(ThisBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("A2:E" & importBottomRow).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
             
        ' If v.3.05 or earlier read from "G"; otherwise read from "F"
             
        If TSVer < 3.06 Then
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("G2:G" & importBottomRow).Select
             Selection.Copy
        Else
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("F2:F" & importBottomRow).Select
             Selection.Copy
        End If
             Workbooks(ThisBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("F2:F" & importBottomRow).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
             
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("L2:M" & importBottomRow).Select
             Selection.Copy
             Workbooks(ThisBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("L2:M" & importBottomRow).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
             
        Range("A2").Select                      ' Place cursor back at home
  
        End If
        
        ' Close the Source Workbook
        TS_Import_Close_Workbook
       
        ' Let user know we're done
        Workbooks(ThisBookName).Activate 'Select original workbook
        Worksheets(ThisSheetName).Activate 'And original worksheet
        result = MsgBox("Successfully imported data from file" & Chr(13) & _
                    """" & DataBookName & """.", vbInformation)
    
' Let's turn calculation on again
    Application.Calculation = xlCalculationAutomatic
 
    
'If Timesheet was originally protected, reprotect
If tsProtected Then
    Sheets("Timesheet").Protect
End If
  
    
    End If
End Sub

Private Sub TS_Import_Close_Workbook()
'
' Closes the data workbook
'
    If DataBookName <> "" Then
        Workbooks(DataBookName).Worksheets(1).Range("A1").Copy 'Copy a single cell to empty large buffer
        Workbooks(DataBookName).Close SaveChanges:=False
    End If
End Sub

Function CCTrim(ChargeCode)
' Need to remove the trailing spaces from the charge codes / work packages
' without removing any of the spaces in the middle of the code

' RTRIM does just that, but is available only in VBA and not as an Excel function
' TRIM will remove multiple internal spaces, corrupting AD          MET
' CLEAN will not remove trailing spaces

CCTrim = RTrim(ChargeCode)

End Function