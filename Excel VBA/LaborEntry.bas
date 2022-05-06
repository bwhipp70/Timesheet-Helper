' Timesheet Helper Comments
'
' 3.20 - 3 January 2021 - Added a LE_Enter_Labor_Flex410()
' 4.00 - 13 October 2021 - Moved common data from IE porton, added Browser Driver support
'
'**********************************************************

'Macro Module: LaborEntry
'Last Updated: 2021-10-12 WRH

'This macro module is used in UpTEMPO

'Purpose: Manage public variables and call labor entry macros
' for different labor work schedules and browser/driver combinations

'Recent changes (in reverse chronological order):
' 2021-10-12 WRH: Updated to support Edge and Chrome using SeleniumBasic (FOSS)
' 2020-12-28 WRH: Updated to support new sheet "Labor_Flex410" for new 4x10 schedule in TEMPO

Option Explicit
'
' URLs for TEMPO
'
Public URL_TEMPO As String
Public Suffix_Shell_Home As String
Public Suffix_Time_Entry As String
Public URL_LoggedOff As String
Public URL_Authentication As String
'
' Browser driver
'
Public theBrowserDriver As String
'
' Defaults
'
Public Const Default_URL_Authenticaton = "https://auth.p.external.lmco.com/idp/SSO.saml2"
Public Const Default_URL_TEMPO = "https://tempo.external.lmco.com/fiori"
Public Const Default_Suffix_Shell_Home = "#Shell-home"
Public Const Default_Suffix_Time_Entry = "#ZTPOTIMESHEET3-record"
Public Const Default_URL_LoggedOff = "https://tempo.external.lmco.com/sap/public/bc/icf/logoff"
'
'If desired screen does not appear in TimeOut seconds, abort the operation
'
Public timeout
Public Const DefaultTimeOut = 30
'
'Default delays in seconds
'
Public NoDelay As Integer
Public Const DefaultNoDelay = 0
Public SingleDelay As Integer
Public Const DefaultSingleDelay = 1
Public DoubleDelay As Integer
Public Const DefaultDoubleDelay = 2
'
'Temporarily store user User IDs and passwords
'
Public RACFUserID
Public RACFPassword
'
'Enter all labor, or just rows with non-zero hours?
'
Public AllLaborX
'
'Show dialog when Labor Entry is completed?
'
Public CompletedDialogX

Sub LE_Enter_Labor_Flex410()
    Call LE_EnterLabor(Labor_Flex410_ShName)
End Sub

Sub LE_Enter_Labor_Flex980()
    Call LE_EnterLabor(Labor_Flex980_ShName)
End Sub

Sub LE_Enter_Labor_Flex980_2weeks()
    Call LE_EnterLabor(Labor_Flex980_2weeks_ShName)
End Sub

Sub LE_EnterLabor(CallingSheet)
    Dim result As Integer
    If Range("Educational_Mode") = "Off" Then
        result = MsgBox("This feature has been disabled.", vbExclamation)
        End
    End If
    result = MsgBox("This is for Educational Purposes ONLY!" & Chr(13) & _
                    "DO NOT USE THIS TO ENTER YOUR OFFICIAL TIMECARD!", vbExclamation)
    If Not CheckAllHash Then
        result = MsgBox("The Macro Code has been corrupted!" & Chr(13) & _
                        "Please grab the released file, import, and try again.", vbExclamation)
        End
    End If
    Call LE_GetUserValues(CallingSheet)
    If theBrowserDriver = IE_BrowserDriver Then
        Call IE_EnterLabor(CallingSheet)
    ElseIf theBrowserDriver = SB_Edge_BrowserDriver Then
        Call SB_EnterLabor(CallingSheet)
    ElseIf theBrowserDriver = SB_Chrome_BrowserDriver Then
        Call SB_EnterLabor(CallingSheet)
    Else
        Call Debug_Warn_User("LE_EnterLabor", "Unknown Browser Driver")
    End If
End Sub

Sub LE_GetUserValues(CallingSheet)
' Gets common user values from Instructions page
'
' 4.00 - Added the Browser Driver

Dim result
    URL_TEMPO = LE_GetSetValue("TEMPO_URL", Default_URL_TEMPO)
    Suffix_Shell_Home = LE_GetSetValue("TEMPO_ShellHome_Suffix", Default_Suffix_Shell_Home)
    Suffix_Time_Entry = LE_GetSetValue("TEMPO_TimeEntry_Suffix", Default_Suffix_Time_Entry)
    URL_LoggedOff = LE_GetSetValue("TEMPO_LoggedOff_URL", Default_URL_LoggedOff)
    URL_Authentication = Default_URL_Authenticaton
    AllLaborX = Range("AllLabor_X").Value
    CompletedDialogX = Range("CompletedDialog_X").Value
    theBrowserDriver = Range("BrowserDriver").Value
    timeout = Range("Timeout_Delay").Value
    If timeout < 1 Then
        timeout = DefaultTimeOut
        Range("Timeout_Delay") = timeout
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

Function LE_GetSetValue(rangeName, defaultValue)
' Check for user value, if blank, sets to default
' then return the user value
'
    If Range(rangeName).Value = "" Then
        Range(rangeName).Value = defaultValue
    End If
    LE_GetSetValue = Range(rangeName).Value
End Function

'Code Module SHA-512
'''1a97961ffa86e022dd1bc4c528edb602fb17b17809c12fb7af1f0f6ae1b1a9c442427775ecdf67dbffe3edc3b1a41d497e5a73016ac779fb9e0cf8a2560c9611