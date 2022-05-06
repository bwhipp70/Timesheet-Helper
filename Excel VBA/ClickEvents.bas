'Macro Module: ClickEvents
'Last Updated: 2021-01-04 WRH

'This macro module is used in UpTEMPO - introduced in version 1.1r1

'Purpose: Handle button clicks from UpTEMPO
' The "_Click" subroutines here should only call other macros
' Subroutine naming convention: [Sheet name]_[Button name]_Click()

'Recent changes (in reverse chronological order):
' 2021-01-04 WRH: Added public constant Debug_Warn
' 2020-12-31 WRH: Updated debug and execution logic (added Perform_Click constant)
' 2020-12-28 WRH: Created for cleaner definition of button actions

Option Explicit

' Debug options to verify button functions

Private Const Debug_Click As Boolean = False    'set to True for debug, False for normal operation
Private Const Perform_Click As Boolean = True   'set to False for debug, True for normal operation
Public Const Debug_Warn As Boolean = False      'set to True for debug, False for normal operation

Private Sub Display_Button_Sub(theSub As String)
    ' Display a MsgBox if debugging is enabled
    Dim result As VbMsgBoxResult
    Dim thePrompt As String
    Dim theButtons As Long
    
    thePrompt = "Debugging is enabled for ClickEvents" & Chr(10) & Chr(10)
    theButtons = vbInformation
    If Perform_Click Then
        ' tell user what subroutine the button is about call and allow user to cancel
        thePrompt = thePrompt & "This button is about to call"
        theButtons = theButtons + vbOKCancel
    Else
        ' tell user what subroutine the button normally calls
        ' but not calling the subroutine, so no need to cancel
        thePrompt = thePrompt & "This button calls"
        theButtons = theButtons + vbOKOnly
    End If
    thePrompt = thePrompt & " subroutine: " & theSub

    result = MsgBox(Prompt:=thePrompt, _
                    Buttons:=theButtons, _
                    Title:="Display_Button_Sub")
    
    If result = vbCancel Then
        ' User canceled - end the macro!
        End
    End If
End Sub

Sub Warn_Button_Debug()
    ' Display a warning if Debug_Click is true or Perform_Click is false
    Dim result As VbMsgBoxResult
    Dim thePrompt As String
    
    If (Debug_Click) Or (Not Perform_Click) Or (Debug_Warn) Then
        thePrompt = "Macro Constant Settings Warning:" & Chr(10) & Chr(10)
        If Debug_Click Then
            thePrompt = thePrompt & "Private Const Debug_Click is True, should be False" & Chr(10)
        End If
        If Not Perform_Click Then
            thePrompt = thePrompt & "Private Const Perform_Click is False, should be True" & Chr(10)
        End If
        If Debug_Warn Then
            thePrompt = thePrompt & "Public Const Debug_Warn is True, should be False" & Chr(10)
        End If
        thePrompt = thePrompt & Chr(10) & "Please fix in macro module ClickEvents!"
        
        result = MsgBox(Prompt:=thePrompt, _
                        Buttons:=vbExclamation + vbOKOnly, _
                        Title:="Warn_Button_Debug")
    End If
End Sub

' Buttons on Worksheet: Instructions

Sub Instructions_Update_Click()
    If Debug_Click Then Call Display_Button_Sub("Instructions_Update_Click")
    If Perform_Click Then Call Update_Work_Schedule_Selection     'in Module: LaborRoutines
End Sub

Sub Instructions_Import_Click()
    If Debug_Click Then Call Display_Button_Sub("Instructions_Import_Click")
    If Perform_Click Then Call Import_From_Other_Workbook         'in Module: LaborRoutines
End Sub

Sub Instructions_Copy_Click()
    If Debug_Click Then Call Display_Button_Sub("Instructions_Copy_Click")
    If Perform_Click Then Call Copy_From_Other_Work_Schedule      'in Module: LaborRoutines
End Sub

Sub Instructions_New_Week_Click()
    If Debug_Click Then Call Display_Button_Sub("Instructions_New_Week_Click")
    If Perform_Click Then Call ClearLaborHours                    'in Module: LaborRoutines
End Sub

' Buttons on Worksheet: Labor_Flex410

Sub Labor_Flex410_Enter_Labor_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex410_Enter_Labor_Click")
    If Perform_Click Then Call LE_Enter_Labor_Flex410              'in Module: LaborEntry
End Sub

Sub Labor_Flex410_Import_Labor_Log_from_SAP_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex410_Import_Labor_Log_from_SAP_Click")
    If Perform_Click Then Call Import_Labor_Log_SAP_Flex410        'in Module: ImportLaborLog
End Sub

Sub Labor_Flex410_Sort_Memo_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex410_Sort_Memo_Click")
    If Perform_Click Then Call Labor_Sort_Memo_Flex410             'in Module: LaborRoutines
End Sub

Sub Labor_Flex410_Sort_Charge_Object_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex410_Sort_Charge_Object_Click")
    If Perform_Click Then Call Labor_Sort_Workpackage_Flex410      'in Module: LaborRoutines
End Sub

Sub Labor_Flex410_Clear_Prior_Weeks_Total_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex410_Clear_Prior_Weeks_Total_Click")
    If Perform_Click Then Call ClearPriorWeeks_Flex410             'in Module: LaborRoutines
End Sub

' Buttons on Worksheet: Labor_Flex980

Sub Labor_Flex980_Enter_Labor_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_Enter_Labor_Click")
    If Perform_Click Then Call LE_Enter_Labor_Flex980             'in Module: LaborEntry
End Sub

Sub Labor_Flex980_Import_Labor_Log_from_SAP_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_Import_Labor_Log_from_SAP_Click")
    If Perform_Click Then Call Import_Labor_Log_SAP_Flex980       'in Module: ImportLaborLog
End Sub

Sub Labor_Flex980_Sort_Memo_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_Sort_Memo_Click")
    If Perform_Click Then Call Labor_Sort_Memo_Flex980            'in Module: LaborRoutines
End Sub

Sub Labor_Flex980_Sort_Charge_Object_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_Sort_Charge_Object_Click")
    If Perform_Click Then Call Labor_Sort_Workpackage_Flex980     'in Module: LaborRoutines
End Sub

' Buttons on Worksheet: Labor_Flex980_2weeks

Sub Labor_Flex980_2weeks_Enter_Labor_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_2weeks_Enter_Labor_Click")
    If Perform_Click Then Call LE_Enter_Labor_Flex980_2weeks          'in Module: LaborEntry
End Sub

Sub Labor_Flex980_2weeks_Import_Labor_Log_from_SAP_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_2weeks_Import_Labor_Log_from_SAP_Click")
    If Perform_Click Then Call Import_Labor_Log_SAP_Flex980_2weeks    'in Module: ImportLaborLog
End Sub

Sub Labor_Flex980_2weeks_Sort_Memo_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_2weeks_Sort_Memo_Click")
    If Perform_Click Then Call Labor_Sort_Memo_Flex980_2weeks         'in Module: LaborRoutines
End Sub

Sub Labor_Flex980_2weeks_Sort_Charge_Object_Click()
    If Debug_Click Then Call Display_Button_Sub("Labor_Flex980_2weeks_Sort_Charge_Object_Click")
    If Perform_Click Then Call Labor_Sort_Workpackage_Flex980_2weeks  'in Module: LaborRoutines
End Sub

' Buttons on Worksheet: Simple Labor Adjustment

Sub Simple_Labor_Adjustment_Clear_Click()
    If Debug_Click Then Call Display_Button_Sub("Simple_Labor_Adjustment_Clear_Click")
    If Perform_Click Then Call ClearLaborAdjustment               'in Module: LaborRoutines
End Sub

'Code Module SHA-512
'''8bcffa01d3e308f0b2849b0e696509c78cc2c4804e291d746812988af858842478beb4055ad320c2c8b409cb2fc0f7c0bcd992070ce0802c2a07ffdc5c0aeecb