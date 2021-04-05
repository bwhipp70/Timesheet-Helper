' Timesheet Helper Comments
'
' 3.20 - 3 January 2021 - Added a LE_Enter_Labor_Flex410()
'
'**********************************************************

Option Explicit
'
'If desired screen does not appear in TimeOut seconds, abort the operation
Public TimeOut
Public Const DefaultTimeOut = 30
'
'Default delays in seconds
Public NoDelay As Integer
Public Const DefaultNoDelay = 0
Public SingleDelay As Integer
Public Const DefaultSingleDelay = 1
Public DoubleDelay As Integer
Public Const DefaultDoubleDelay = 2
'
'Temporarily store user User IDs and passwords
Public RACFUserID
Public RACFPassword
'Enter all labor, or just rows with non-zero hours?
Public AllLaborX
'Show dialog when Labor Entry is completed?
Public CompletedDialogX

Sub LE_Enter_Labor_Flex980()
    Call IE_EnterLabor(Labor_Flex980_ShName)
End Sub
Sub LE_Enter_Labor_Flex980_2weeks()
    Call IE_EnterLabor(Labor_Flex980_2weeks_ShName)
End Sub
Sub LE_Enter_Labor_Flex410()
    Call IE_EnterLabor(Labor_Flex410_ShName)
End Sub
Sub LE_GetUserValues(CallingSheet)
' Gets common user values from Instructions page
'
Dim result
    TimeOut = Range("Timeout_Delay").Value
    If TimeOut < 1 Then
        TimeOut = DefaultTimeOut
        Range("Timeout_Delay") = TimeOut
    End If
    NoDelay = DefaultNoDelay  'no user entry for NoDelay
    SingleDelay = Range("Single_Delay").Value
    If SingleDelay < 1 Then
        SingleDelay = DefaultSingleDelay
        Range("Single_Delay") = SingleDelay
    End If
    DoubleDelay = Range("Double_Delay").Value
    If DoubleDelay < 1 Then
        DoubleDelay = DefaultDoubleDelay
        Range("Double_Delay") = DoubleDelay
    End If
End Sub
