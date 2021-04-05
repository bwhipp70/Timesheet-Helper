'Subroutines for use by all modules
Option Explicit
'
Public HatchCols As Long
Public HatchRows As Long
'
'WinAPI functions
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal _
 hwnd As Long) As Long

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal _
 lpClassName As Any, ByVal lpWindowName As Any) As Long

Private Declare PtrSafe Function GetTopWindow Lib "user32" (ByVal _
 hwnd As Long) As Long

Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal _
 hwnd As Long) As Long

Private Declare PtrSafe Function OpenIcon Lib "user32" (ByVal _
 hwnd As Long) As Long
'
'
Sub Set_Row_Col_Hatch_Ranges()
Dim sheetName
    sheetName = ActiveSheet.Name
    If sheetName = Instructions_ShName Then
        HatchCols = 12
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

Public Sub IEFrameToTop()
 Dim THandle As Long
 
 THandle = FindWindow("IEFrame", vbEmpty)
 
 If THandle = 0 Then
  MsgBox "Could not find window.", vbOKOnly
 Else
  BringWindowToTop THandle
 End If
End Sub
Sub Workbook_Activate()
' Activates the Excel workbook in which this macro is running
' Note: ActiveWorkbook is the workbook in the active (frontmost) window in Excel
'       ThisWorkbook is the workbook in which this macro is running
Dim isActive As Boolean
    isActive = True
    If Application.ActiveWorkbook Is Nothing Then
        isActive = False
    ElseIf Not (Application.ActiveWorkbook.Name = Application.ThisWorkbook.Name) Then
        isActive = False
    End If
    If Not isActive Then
        Workbooks(Application.ThisWorkbook.Name).Activate
    End If
End Sub
Sub Excel_Activate()
' Activates Excel - brings Excel window (and message box, if open) to front
Dim appWindowTitle As String
Dim THandle As Long
        
' Sometimes activating Excel doesn't bring the window to the front - it seems
' like Internet Explorer is holding the focus.  So we tab to the next window first:

    'workaround for SendKeys bug (see https://support.microsoft.com/en-us/kb/179987)
    DoEvents
    Get_Keyboard_States
    Application.SendKeys ("%{TAB}")     'Alt+TAB
    Set_Keyboard_States
    DoEvents
    
' This no longer works in Excel 2013 (also started failing in Excel 2010 on 9/21/2015):
'    AppActivate ("Microsoft Excel")
' This often works in Excel 2013, but sometimes fails:
'    AppActivate ("Excel")
' So we use this new method:
' (note that without an object qualifier, Application represents the entire Microsoft Excel application)
    appWindowTitle = Application.Caption    ' Get the current Excel window title
    AppActivate (appWindowTitle)            ' Use AppActivate with full window title
' Debug: print appWindowTitle to Immediate window
    'Debug.Print "Excel_Activate", appWindowTitle
    DoEvents
    THandle = GetTopWindow(vbEmpty)
    'Debug.Print "GetTopWindow", THandle
    
' Also use WinAPI functions to bring window to top (in case window is minimized)
    THandle = FindWindow(vbEmpty, appWindowTitle)
    'Debug.Print "FindWindow", THandle
    If Not (THandle = 0) Then
        If IsIconic(THandle) <> 0 Then
            ' Open iconic (minimized) window
            OpenIcon THandle
        End If
        ' Bring window to top
        BringWindowToTop THandle
        DoEvents
        THandle = GetTopWindow(vbEmpty)
        'Debug.Print "GetTopWindow (inside If)", THandle
    End If
End Sub
Sub Set_Calculation(turnOn)
    If turnOn Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub
'
' Functions for conditional formatting
'
Public Function IsValidChargeObject(theString As String) As Boolean
    Dim result As Boolean
    Dim theLength
    Dim i As Integer
    
    result = False
    theLength = Len(theString)
    If (theLength = 0) Or (theLength = 1) Or (theLength = 2) Or (theLength = 12) Then
        result = True
    ElseIf (theLength = 15) Then
        If UCase(Left(theString, 1)) = "P" Then
            result = True
            For i = 2 To 15
                If (Mid(theString, i, 1) < "0") Or _
                   (Mid(theString, i, 1) > "9") Then
                    result = False
                End If
            Next i
        End If
    End If
    IsValidChargeObject = result
End Function
'
Private Sub TestIVCO()
    Dim s0 As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String
    Dim s5 As String
    Dim s6 As String
    s0 = ""
    s1 = "A"
    s2 = "BC"
    s3 = "DEF"
    s4 = "GHIJKLMNOPQR"
    s5 = "p00100056782345"
    s6 = "P10000000o00234"
    MsgBox s0 & " = " & IsValidChargeObject(s0) & Chr(10) & _
        s1 & " = " & IsValidChargeObject(s1) & Chr(10) & _
        s2 & " = " & IsValidChargeObject(s2) & Chr(10) & _
        s3 & " = " & IsValidChargeObject(s3) & Chr(10) & _
        s4 & " = " & IsValidChargeObject(s4) & Chr(10) & _
        s5 & " = " & IsValidChargeObject(s5) & Chr(10) & _
        s6 & " = " & IsValidChargeObject(s6)
End Sub

