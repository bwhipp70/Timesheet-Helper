'*************************************************************
'Macro Warning Routines
'WRH 2011-01-22
'
' Added Disable Cut and Paste and Drag and Drop
' BMW, Ver 3.22
'
Option Explicit
'
'This is the name of the worksheet containing the macro warning information
Private Const Macro_Warning_Sheet_Name = "Macro Warning"
'
'This is the named cell on the macro warning sheet where the active sheet is saved
' before showing the macro warning sheet
Private Const MW_Working_Sheet_Ref = "MW_Working_Sheet"
'
'This is the named cell on the instructions sheet where an "X" is placed to enable
' the macro warning routines.  (Actually, any character will work: cell value <> "")
Private Const MW_Enabled_Ref = "MacroWarning_X"
'
'This variable stores whether we're closing the workbook
Private isClosing As Boolean
'
'This variable stores whether we're making expected changes to the workbook
Private ignoreChanges As Boolean
'
'Event Handler called by Excel before closing the workbook
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Let other events know that the workbook is closing now
    isClosing = True
    'Problem is, if user cancels the close this variable is still true, and saves won't work correctly
    'Workaround is sheet change event handlers below
End Sub
'
'Event Handler called by Excel before saving the workbook
'
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim fileSaveName
    'Store the active sheet name and select the Macro Warning sheet
    Show_Macro_Warning
    'Excel 2007 and before do not have an AfterSave event
    'The following idea was found at http://www.vbforums.com/showthread.php?t=524692
    ' See response #16 from DrillRaider on Sep 24th 2010 at 12:38 PM
    'Set up timed call to AfterSave routine (will run after normal save is complete!)
    'But only do this if we're saving without closing, otherwise Excel will re-open the workbook!
    If Not isClosing Then
        Application.OnTime Now, "ThisWorkbook.AfterSave"
    End If
    'Allow the normal save routines to work
End Sub
'
'Event Handler called by Excel after saving the workbook
'
Private Sub AfterSave()
    'Change back to the originally active sheet
    Hide_Macro_Warning
End Sub
'
'Event Handler called by Excel when opening the workbook (if/when macros are enabled)
'
Private Sub Workbook_Open()
    'Restore the active sheet and hide the Macro Warning sheet
    Hide_Macro_Warning
    'We're not closing the workbook - initialize private variable
    isClosing = False
    'We're not saving the workbook - initialize private variable
    ignoreChanges = False
End Sub
'
'Event handlers for sheet changes and window changes
' If we see activity here, we're not closing the workbook!
'
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Not ignoreChanges Then
        isClosing = False
    End If
End Sub
Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
' Removing the flag being set, causes a VBA Error Code 400 when another Workbook is open
' V3.15
'    isClosing = False
End Sub
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not ignoreChanges Then
        isClosing = False
    End If
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'* Added disable cut and paste
'* From https://www.mrexcel.com/board/threads/disable-cut-function.694926/
'* Ver 3.22
    isClosing = False
    
' Can't disable Drag and Drop, will break CLean for Distribution
'    Application.CellDragAndDrop = False
    
    Select Case Application.CutCopyMode
        Case 2
            MsgBox "The 'Cut & Paste' functionality is not available in this file." & vbCr & vbCr & _
                "Please use the 'Copy & Paste' funtionality only and manually clear any data that require moving", vbExclamation, "Function Unavailable"
            Application.CutCopyMode = False 'clear clipboard and cancel cut
        Case Else
    End Select

End Sub
Private Sub Workbook_Deactivate()
     Application.CellDragAndDrop = True
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    isClosing = False
End Sub
Private Sub Workbook_WindowResize(ByVal Wn As Window)
    isClosing = False
End Sub
'
'Function to check whether Macro Warning is enabled
'
Private Function MacroWarning_Enabled() As Boolean
    MacroWarning_Enabled = Range(MW_Enabled_Ref).Value <> ""
End Function
'
'Subroutine to store the active sheet name and show the Macro Warning sheet
'
Private Sub Show_Macro_Warning()
Dim wasSaved As Boolean
    If MacroWarning_Enabled Then
        'Disable screen updating
        Application.ScreenUpdating = False
        'Store current saved state for this workbook
        wasSaved = Me.Saved
        'Ignore expected sheet activations and changes
        ignoreChanges = True
        'Show the macro warning sheet
        Me.Sheets(Macro_Warning_Sheet_Name).Visible = True
        'Save the active sheet name
        Me.Sheets(Macro_Warning_Sheet_Name).Range(MW_Working_Sheet_Ref).Value = Me.ActiveSheet.Name
        'Select the macro warning sheet
        Me.Sheets(Macro_Warning_Sheet_Name).Select
        'Select the top row on the macro warning sheet to hide the cursor (top row is hidden)
        Me.Sheets(Macro_Warning_Sheet_Name).Rows(1).Select
        'Don't ignore sheet activations and changes
        ignoreChanges = False
        'Showing (un-hiding) the macro warning sheet counts as a change to the workbook.
        ' If it was already saved, reset the saved flag to ignore this change
        If wasSaved Then
            Me.Saved = True
        End If
    End If
End Sub
'
'Subroutine to return to the active sheet and hide the Macro Warning sheet
'
Private Sub Hide_Macro_Warning()
Dim wasSaved As Boolean
    If MacroWarning_Enabled Then
        'Store current saved state for this workbook
        wasSaved = Me.Saved
        'Select the saved sheet (this was the active sheet before showing the macro warning sheet)
        Me.Sheets(Range(MW_Working_Sheet_Ref).Value).Select
        'Hide the macro warning sheet
        Me.Sheets(Macro_Warning_Sheet_Name).Visible = False
        'Hiding the macro warning sheet counts as a change to the workbook.
        ' If it was already saved, reset the saved flag to ignore this change
        If wasSaved Then
            Me.Saved = True
        End If
        'Enable screen updating
        Application.ScreenUpdating = True
    End If
End Sub
'
'End of Macro Warning Routines
'*************************************************************
