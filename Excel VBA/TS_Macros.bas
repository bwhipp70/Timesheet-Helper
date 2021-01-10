' *******************************************
' Timesheet Macros
'
' Version:
' 3.00 - Brian Whipp, based on UpTEMPO 1.0a3 (2016-09-13)
' 3.01 - Brian Whipp, included updates from UpTEMPO 1.0a4 (2016-09-23)
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

'3. Sheet Labor_Flex980
'   Changed N5 to =BH10
'4. Sheet Labor_Flex980_2weeks
'   Changed T5 to =BH10
'5. WP #'s - Column I was originally sorted, removing duplicates and blanks.  This slows things down considerably.
'   Name Manager - WP_List =OFFSET('WP #''s'!$A$2,0,0,COUNTA('WP #''s'!$A$2:$A$150))
'   Name Manager - WP_List_Unique_alpha =OFFSET('WP #''s'!$I$2, 0, 0, COUNT(IF('WP #''s'!$I$2:$I$149="", "", 1)), 1)
'   WP #'s, Column I, cell I:2 {=IFERROR(INDEX(WP_List, MATCH(0, IF(MAX(NOT(COUNTIF($I$1:I1, WP_List))*(COUNTIF(WP_List, ">"&WP_List)+1))=(COUNTIF(WP_List, ">"&WP_List)+1), 0, 1), 0)),"")}
'   Timing Values, COlumn I, Baseline = 1.88005
'   Timing Values, WP_List = A2:A150 = 1.71694
'   Timing Values, WP_List_Alpha_Unique = Column I = 2.0026
'   Numbers don't make much sense, performance is noticeably better?!





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
  lastRow = sht.Cells(sht.Rows.Count, "C").End(xlUp).Row + 1
  BottomRow = sht.Cells(sht.Rows.Count, "O").End(xlUp).Row + 1

If TS_MaxRows < lastRow Then
    Sheets("Configuration").Range("E20").Value = lastRow + 2
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
    Range("I28").Value = "X"

' Web Interface Tuning
    Range("C63").Value = "15"               ' Timeout
    Range("C65").Value = "1"                ' Delay
    Range("C67").Value = "2"                ' Double Delay
    
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

    Range("H19").Value = "X"                 ' Import Configuration
    Range("H20").Value = "X"                 ' Import WP #'s
    Range("H21").Value = "X"                 ' Import Timesheet


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

Sheets("TSMasterFormulas").Select

Range("A2:N2").Select
Selection.Copy

Sheets("Timesheet").Select

' Set up "first" row
' Will need to do all cells in the row
    
With Sheets("Timesheet").Range("A" & lastRow)
    .PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With

Sheets("TSMasterFormulas").Select

Range("O2:AR2").Select
Selection.Copy

Sheets("Timesheet").Select

' Set up "first" row
' Will need to do all cells in the row
    
With Sheets("Timesheet").Range("O" & lastRow)
    .PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End With


If TS_MaxRows < 2 Then
    TS_MaxRows = 2
End If

If TS_MaxRows < lastRow Then
    TS_MaxRows = lastRow
End If

If Resize Then
    ' Start from LastRow to BottomRow
    tempRange = "A" & lastRow & ":AR" & TS_MaxRows
    Worksheets("Timesheet").Range(tempRange).FillDown
Else
    Worksheets("Timesheet").Range("A2:AR" & TS_MaxDefaultRows).FillDown
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
    
    Range("B11:B150").Select
    Selection.ClearContents
   
    Range("C2:G150").Select
    Selection.ClearContents
    
    Range("A11:A150").Value = "_blank_"
    
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
Dim Continue_Import As Boolean
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

Continue_Import = False

testString = ""

' Determine which sheets to import
configSh = False
wpSh = False
timesheetSh = False

Sheets("Configuration").Select

If (Range("H19").Value <> "") Then
    configSh = True
End If

If (Range("H20").Value <> "") Then
    wpSh = True
End If

If (Range("H21").Value <> "") Then
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
   
        ' Put in code to transfer data here
        ' Configuration and WP #'s will be version based
        ' Transfer in appropriate data
                       
        ' this needs help, having trouble detecting a sheet existence
        
        SheetExists = False
        
        ' Make sure it is a valid workbook
        ' If not, trap the error and gracefully inform the user
        
        On Error Resume Next
        Workbooks(DataBookName).Activate
        testString = Worksheets("Change History").Range("A1")
        On Error GoTo 0
        If testString = "Rev History" Then
            SheetExists = True
        Else
            SheetExists = False
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
            TSVer = Sheets("Configuration").Range("X2").Value
        End If
           
'        MsgBox "SheetExists = " & SheetExists & "; V208/209/300 = " & v208 & ";" & v209 & ";" & v300
'        MsgBox "TSVer = " & TSVer
        
'       Needs to be V2.08 or newer to import.

        If (Not v208) And (Not v209) And (Not v300) Then
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
         
           Workbooks(ThisBookName).Activate
           Sheets("Configuration").Select
           Sheets("Configuration").Range("A2").Value = theValues(0)
           Sheets("Configuration").Range("C2").Value = theValues(1)
           Sheets("Configuration").Range("E2").Value = theValues(2)
           Sheets("Configuration").Range("G2").Value = theValues(3)
           Sheets("Configuration").Range("I2").Value = theValues(4)
           Sheets("Configuration").Range("K2").Value = theValues(5)
         
           If v300 Then
             Sheets("Configuration").Range("A20").Value = theValues(6)
             Sheets("Configuration").Range("E20").Value = theValues(7)
           End If
     
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
            End If
          End If
        
        End If
        
' Import the Timesheet entries

        If timesheetSh Then
        ' Copy everything down to last row
        ' Columns A-E, G, L-M
        ' Determine number of total rows and used rows in source sheet
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
        ' Ctrl + Shift + End
             importBottomRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "O").End(xlUp).Row + 1
             
        ' Set up destination sheet length
             Workbooks(ThisBookName).Activate
             Sheets("Configuration").Select
             Range("E20").Value = importBottomRow - 1
             Call TS_UpdateMaxRows
             
        ' Copy data
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("A2:E" & importBottomRow).Select
             Selection.Copy
             Workbooks(ThisBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("A2:E" & importBottomRow).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
             
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("G2:G" & importBottomRow).Select
             Selection.Copy
             Workbooks(ThisBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("G2:G" & importBottomRow).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
             
             Workbooks(DataBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("L2:M" & importBottomRow).Select
             Selection.Copy
             Workbooks(ThisBookName).Activate
             Sheets("Timesheet").Select
             ActiveSheet.Range("L2:M" & importBottomRow).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
             
  
        End If
        
        ' Close the Source Workbook
        TS_Import_Close_Workbook
       
        ' Let user know we're done
        Workbooks(ThisBookName).Activate 'Select original workbook
        Worksheets(ThisSheetName).Activate 'And original worksheet
        result = MsgBox("Successfully imported data from file" & Chr(13) & _
                    """" & DataBookName & """.", vbInformation)
    
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
