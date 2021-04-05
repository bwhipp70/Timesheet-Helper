' Timesheet Helper Comments
'
' 3.20 - 3 January 2021 - changes to IE_EnterLabor to accomodate new Flex410 sheet
'
'
'***********************************************************************************
Option Explicit
'
' To get ShellWindows: go to the Tools menu at the top of the VBA editor,
'  click on References in the menu, and scroll down the long list to find
'  the “Microsoft Internet Controls” reference. Click the checkbox to the
'  left of it, and then click OK.
'
' URLs for TEMPO
'
Public URL_TEMPO As String
Public Suffix_Shell_Home As String
Public Suffix_Time_Entry As String
Public URL_LoggedOff As String

Public Const Default_URL_TEMPO = "https://tempo.external.lmco.com/fiori" 'Updated 3.18
Public Const Default_Suffix_Shell_Home = "#Shell-home"
Public Const Default_Suffix_Time_Entry = "#ZTPOTIMESHEET3-record" 'Updated 3.18
Public Const Default_URL_LoggedOff = "https://tempo.external.lmco.com/sap/public/bc/icf/logoff" 'Updated 3.18
'
' Keep track of whether TEMPO time entry sheet has "WD Job" field
'  default is empty string (need to check), "Y" if yes, and "N" if no
'
Dim Found_WD_Job As String
'
'WinAPI functions
#If Win64 Then
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal _
 hwnd As LongPtr) As Long
 
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal _
 hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
#Else
Private Declare Function BringWindowToTop Lib "user32" (ByVal _
 hwnd As Long) As Long
 
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal _
 hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
#End If
'
'
Sub IE_EnterLabor(CallingSheet)
' Logs in to TEMPO and enters labor from Labor Worksheet
'
    Dim objIE As Object
    Dim objDocument As Object
    Dim result As Integer
    Dim FirstLaborRow, LastLaborRow
    Dim colTotalHours
    Dim pageStr As String
    Dim i As Integer
    Dim MMDDYYYYstr, WEdate, datestr
    Dim iRow, iPage
    Dim iEntries As Integer
    Dim LastPage As Boolean
    Dim LastEntryRow
    Dim ErrorText, ErrorCode
    Dim theHours(7) As String
    Dim entriesMatch As Boolean
    Dim loopCount As Integer
    Const loopLimit As Integer = 3

    Call IE_GetUserValues(CallingSheet)
    
    'Find or open TEMPO in Internet Explorer
    Set objIE = IE_Find_Or_Open_TEMPO()
    Call IE_Wait_Until_Done(objIE)
    
    'check for TEMPO Logged Off page
    If IE_WhatsRunning = "TEMPO Logged Off Page" Then
        objIE.Navigate URL_TEMPO
        Call IE_Wait(DoubleDelay)
        Call IE_Wait_Until_Done(objIE)
    End If
    
    'check for TEMPO Login page
    If IE_WhatsRunning = "TEMPO Login Page" Then
        objIE.Quit  'close this IE window
        Call IE_Wait(DoubleDelay)
        'Open new TEMPO window in Internet Explorer
        Set objIE = IE_Find_Or_Open_TEMPO()
        Call IE_Wait_Until_Done(objIE)
    End If
    
    'go to TEMPO Time Entry page
    If IE_WhatsRunning = "TEMPO Welcome Page" Then
        IE_TimeEntry_TEMPO
    ElseIf IE_WhatsRunning = "TEMPO Other Page" Then
        IE_TimeEntry_TEMPO
    End If
        
    'check for Attendance & Labor Input page and input labor
    If IE_WhatsRunning = "TEMPO Time Entry Page" Then
        'Get the current web browser session
        Set objIE = IE_Find_TEMPO()
        'Bring it to the front
        If Not IE_Activate(objIE) Then
            'could not find TEMPO Time Sheet
            Excel_Activate
            result = MsgBox("Unable to find the Internet Explorer window and tab containing the TEMPO Time Sheet.", vbExclamation)
            IE_Finish
        End If
        'Get the document object
        Set objDocument = objIE.Document
        'configure first and last labor rows for the selected sheet
        If CallingSheet = Labor_Flex980_ShName Then
            FirstLaborRow = FirstLaborRow_Flex980
            LastLaborRow = LastLaborRow_Flex980
            colTotalHours = 15
        ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
            FirstLaborRow = FirstLaborRow_Flex980_2weeks
            LastLaborRow = LastLaborRow_Flex980_2weeks
            colTotalHours = 21
        ElseIf CallingSheet = Labor_Flex410_ShName Then                   ' Added TSHelper 3.20
            FirstLaborRow = FirstLaborRow_Flex410
            LastLaborRow = LastLaborRow_Flex410
            colTotalHours = 15
        End If
        'check whether week ending date in TEMPO matches this week's date
        If (CallingSheet = Labor_Flex980_ShName) Or _
            (CallingSheet = Labor_Flex980_2weeks_ShName) Or _
            (CallingSheet = Labor_Flex410_ShName) Then                     ' Added TSHelper 3.20
            MMDDYYYYstr = IE_GetWEDate_TEMPO(objIE)
            If MMDDYYYYstr = "" Then
                Excel_Activate
                result = MsgBox("Unable to find Payroll W/E date in TEMPO.", vbExclamation)
                IE_Finish
            End If
            If (CallingSheet = Labor_Flex980_ShName) Then
                WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
            ElseIf (CallingSheet = Labor_Flex980_2weeks_ShName) Then
                WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
            ElseIf (CallingSheet = Labor_Flex410_ShName) Then                ' Added TSHelper 3.20
                WEdate = Sheets(CallingSheet).Range("BH10").Value
            End If
            datestr = Format(WEdate, "mm/dd/yyyy")
            If MMDDYYYYstr <> datestr Then
                Excel_Activate
                Sheets(CallingSheet).Select
                Range("Q2").Select
                result = MsgBox("The week ending (W/E) date in TEMPO (" & MMDDYYYYstr & _
                        ") does not match the Current Week Ending date." & Chr(13) & _
                        "Please verify the week ending dates and try again.", vbExclamation)
                IE_Finish
            End If
        End If
        'perform entries, then verify them
        loopCount = 0
        Do
            'enter days off
            If CallingSheet = Labor_Flex980_ShName Then
                For i = 0 To 7
                    Call IE_EnterDaysOff_TEMPO(objIE, Sheets(CallingSheet).Cells(4, 7 + i).Value, _
                        Sheets(CallingSheet).Cells(5, 7 + i).Value, Sheets(CallingSheet).Cells(8, 7 + i).Value)
                Next i
            ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                For i = 0 To 7
                    Call IE_EnterDaysOff_TEMPO(objIE, Sheets(CallingSheet).Cells(4, 13 + i).Value, _
                        Sheets(CallingSheet).Cells(5, 13 + i).Value, Sheets(CallingSheet).Cells(8, 13 + i).Value)
                Next i
            ElseIf CallingSheet = Labor_Flex410_ShName Then                                 ' Added TS Helper 3.20
                For i = 0 To 6
                    Call IE_EnterDaysOff_TEMPO(objIE, Sheets(CallingSheet).Cells(4, 8 + i).Value, _
                        Sheets(CallingSheet).Cells(5, 8 + i).Value, Sheets(CallingSheet).Cells(8, 8 + i).Value)
                Next i
            End If
            'prepare to enter labor
            iRow = FirstLaborRow
            iEntries = 0
            'Find last row with a labor entry
            LastEntryRow = LastLaborRow
            Do While (LastEntryRow >= FirstLaborRow) And _
               (Sheets(CallingSheet).Cells(LastEntryRow, 3).Value = "")
                LastEntryRow = LastEntryRow - 1
            Loop
            'enter labor
            Do
                'Check whether we're entering all labor
                ' always enter labor lines with non-zero hours (check both total in col 14 and O/T in col 15)
                If (AllLaborX <> "") Or _
                    (Sheets(CallingSheet).Cells(iRow, colTotalHours).Value <> "") Then
                    If CallingSheet = Labor_Flex980_ShName Then
                        theHours(0) = Sheets(CallingSheet).Cells(iRow, 7).Value 'Fri
                        theHours(1) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Sat
                        theHours(2) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Sun
                        theHours(3) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Mon
                        theHours(4) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Tue
                        theHours(5) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Wed
                        theHours(6) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Thu
                        theHours(7) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Fri
                        ' Added in Rev 3.21 - sending the WPM Field
                        Call IE_EnterChargeObj_TEMPO(objIE, iEntries, _
                            Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                            Sheets(CallingSheet).Cells(iRow, 3).Value, _
                            Sheets(CallingSheet).Cells(iRow, 5).Value, _
                            Sheets(CallingSheet).Cells(iRow, 6).Value, _
                            theHours)
                    ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                        theHours(0) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Fri
                        theHours(1) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sat
                        theHours(2) = Sheets(CallingSheet).Cells(iRow, 15).Value 'Sun
                        theHours(3) = Sheets(CallingSheet).Cells(iRow, 16).Value 'Mon
                        theHours(4) = Sheets(CallingSheet).Cells(iRow, 17).Value 'Tue
                        theHours(5) = Sheets(CallingSheet).Cells(iRow, 18).Value 'Wed
                        theHours(6) = Sheets(CallingSheet).Cells(iRow, 19).Value 'Thu
                        theHours(7) = Sheets(CallingSheet).Cells(iRow, 20).Value 'Fri
                        ' Added in Rev 3.21 - sending the WPM Field
                        Call IE_EnterChargeObj_TEMPO(objIE, iEntries, _
                            Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                            Sheets(CallingSheet).Cells(iRow, 3).Value, _
                            Sheets(CallingSheet).Cells(iRow, 5).Value, _
                            Sheets(CallingSheet).Cells(iRow, 6).Value, _
                            theHours)
                    ElseIf CallingSheet = Labor_Flex410_ShName Then                                  ' Added TSHelper 3.20
                        'theHours(0) = Sheets(CallingSheet).Cells(iRow, 7).Value 'Fri - deprecated
                        theHours(0) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Mon
                        theHours(1) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Tue
                        theHours(2) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Wed
                        theHours(3) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Thu
                        theHours(4) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Fri
                        theHours(5) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Sat
                        theHours(6) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sun
                        ' Added in Rev 3.21 - sending the WPM Field
                        Call IE_EnterChargeObj_TEMPO(objIE, iEntries, _
                            Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                            Sheets(CallingSheet).Cells(iRow, 3).Value, _
                            Sheets(CallingSheet).Cells(iRow, 5).Value, _
                            Sheets(CallingSheet).Cells(iRow, 6).Value, _
                            theHours)
                    End If
                    iEntries = iEntries + 1
                End If
                iRow = iRow + 1
            Loop Until iRow > LastEntryRow
            'delete extra labor lines at end
            Call IE_DeleteRows_TEMPO(objIE, iEntries)
            'pause briefly
            Call IE_Wait(1)
            'begin verification of entered data
            entriesMatch = True  'start as true, change to false if any entry does not match
            'verify days off
            
            '*******************
            '* Version 3.16 PTR
            '* Error when on OFF Friday has regular hours
            '* The flag for the day on TIMESHEET Tab is cleared
            '* But TEMPO has the flag set to OFF, therefore a mismatch
            '* Easiest Fix changes to not check for Friday Flag
            '*******************
            
            '************************
            '* Version 3.19 Fix
            '* Removed hack to not check Friday Flag
            '* based on new logic provided by Daniel Eby.
            '* Now will check Friday Flag.
            '************************
            
            If CallingSheet = Labor_Flex980_ShName Then
                For i = 0 To 7 'Was 6 with V3.16
                    If entriesMatch Then
                        entriesMatch = IE_VerifyDaysOff_TEMPO(objIE, Sheets(CallingSheet).Cells(4, 7 + i).Value, _
                            Sheets(CallingSheet).Cells(5, 7 + i).Value, Sheets(CallingSheet).Cells(8, 7 + i).Value)
                    End If
                Next i
            ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                For i = 0 To 7 'Was 6 with V3.16
                    If entriesMatch Then
                        entriesMatch = IE_VerifyDaysOff_TEMPO(objIE, Sheets(CallingSheet).Cells(4, 13 + i).Value, _
                            Sheets(CallingSheet).Cells(5, 13 + i).Value, Sheets(CallingSheet).Cells(8, 13 + i).Value)
                    End If
                Next i
            ElseIf CallingSheet = Labor_Flex410_ShName Then                ' Added TSHelper 3.20
                For i = 0 To 6 'Was 6 with V3.16
                    If entriesMatch Then
                        entriesMatch = IE_VerifyDaysOff_TEMPO(objIE, Sheets(CallingSheet).Cells(4, 8 + i).Value, _
                            Sheets(CallingSheet).Cells(5, 8 + i).Value, Sheets(CallingSheet).Cells(8, 8 + i).Value)
                    End If
                Next i
            End If
            'prepare to verify labor
            iRow = FirstLaborRow
            iEntries = 0
            'verify labor
            Do
                If entriesMatch Then
                    'Check whether we're entering all labor
                    ' always enter labor lines with non-zero hours (check both total in col 14 and O/T in col 15)
                    If (AllLaborX <> "") Or _
                        (Sheets(CallingSheet).Cells(iRow, colTotalHours).Value <> "") Then
                        If CallingSheet = Labor_Flex980_ShName Then
                            theHours(0) = Sheets(CallingSheet).Cells(iRow, 7).Value 'Fri
                            theHours(1) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Sat
                            theHours(2) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Sun
                            theHours(3) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Mon
                            theHours(4) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Tue
                            theHours(5) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Wed
                            theHours(6) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Thu
                            theHours(7) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Fri
                            ' Added in Rev 3.21 - sending the WPM Field
                            entriesMatch = IE_VerifyChargeObj_TEMPO(objIE, iEntries, _
                                Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                                Sheets(CallingSheet).Cells(iRow, 3).Value, _
                                Sheets(CallingSheet).Cells(iRow, 5).Value, _
                                Sheets(CallingSheet).Cells(iRow, 6).Value, _
                                theHours)
                        ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                            theHours(0) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Fri
                            theHours(1) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sat
                            theHours(2) = Sheets(CallingSheet).Cells(iRow, 15).Value 'Sun
                            theHours(3) = Sheets(CallingSheet).Cells(iRow, 16).Value 'Mon
                            theHours(4) = Sheets(CallingSheet).Cells(iRow, 17).Value 'Tue
                            theHours(5) = Sheets(CallingSheet).Cells(iRow, 18).Value 'Wed
                            theHours(6) = Sheets(CallingSheet).Cells(iRow, 19).Value 'Thu
                            theHours(7) = Sheets(CallingSheet).Cells(iRow, 20).Value 'Fri
                            ' Added in Rev 3.21 - sending the WPM Field
                            entriesMatch = IE_VerifyChargeObj_TEMPO(objIE, iEntries, _
                                Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                                Sheets(CallingSheet).Cells(iRow, 3).Value, _
                                Sheets(CallingSheet).Cells(iRow, 5).Value, _
                                Sheets(CallingSheet).Cells(iRow, 6).Value, _
                                theHours)
                        ElseIf CallingSheet = Labor_Flex410_ShName Then                     ' Added TSHelper 3.20
                            'theHours(0) = Sheets(CallingSheet).Cells(iRow, 7).Value 'Fri  deprecated with Flex410
                            theHours(0) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Mon
                            theHours(1) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Tue
                            theHours(2) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Wed
                            theHours(3) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Thu
                            theHours(4) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Fri
                            theHours(5) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Sat
                            theHours(6) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sun
                            ' Added in Rev 3.21 - sending the WPM Field
                            entriesMatch = IE_VerifyChargeObj_TEMPO(objIE, iEntries, _
                                Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                                Sheets(CallingSheet).Cells(iRow, 3).Value, _
                                Sheets(CallingSheet).Cells(iRow, 5).Value, _
                                Sheets(CallingSheet).Cells(iRow, 6).Value, _
                                theHours)
                        End If
                        iEntries = iEntries + 1
                    End If
                End If
                iRow = iRow + 1
            Loop Until iRow > LastEntryRow
            If Not entriesMatch Then
                'increase delays by 1 second each for next try
                NoDelay = NoDelay + 1
                SingleDelay = SingleDelay + 1
                DoubleDelay = DoubleDelay + 1
            End If
            'keep track of number of times through this loop
            loopCount = loopCount + 1
        Loop Until entriesMatch Or (loopCount >= loopLimit)
        'Unable to click the save button at this time, but don't want to automatically click it anyway
        ' - this allows user to review labor before TEMPO combines and rearranges it!
        'pause briefly before returning to Excel
        Call IE_Wait(1)
        Excel_Activate
        If entriesMatch Then
            If (CompletedDialogX <> "") Then
                result = MsgBox("Labor entry completed!" & Chr(10) & Chr(10) & _
                    "Remember to review the labor and click the Save button in TEMPO.", vbInformation)
            End If
        Else
            result = MsgBox("Unable to enter labor correctly!" & Chr(10) & Chr(10) & _
                "Tried " & loopLimit & " times and found at least one incorrect entry each time.", vbExclamation)
        End If
    Else
        'not at Attendance & Labor Input page
        Excel_Activate
        result = MsgBox("Unable to get to TEMPO Time Entry page.", vbExclamation)
        IE_Finish
    End If
    
    Set objIE = Nothing

End Sub
Function IE_GetSetValue(rangeName, defaultValue)
' Check for user value, if blank, sets to default
' then return the user value
'
    If Range(rangeName).Value = "" Then
        Range(rangeName).Value = defaultValue
    End If
    IE_GetSetValue = Range(rangeName).Value
End Function
Sub IE_GetUserValues(CallingSheet)
' Gets user values from Instructions page
'
    URL_TEMPO = IE_GetSetValue("TEMPO_URL", Default_URL_TEMPO)
    Suffix_Shell_Home = IE_GetSetValue("TEMPO_ShellHome_Suffix", Default_Suffix_Shell_Home)
    Suffix_Time_Entry = IE_GetSetValue("TEMPO_TimeEntry_Suffix", Default_Suffix_Time_Entry)
    URL_LoggedOff = IE_GetSetValue("TEMPO_LoggedOff_URL", Default_URL_LoggedOff)
    AllLaborX = Range("AllLabor_X").Value
    CompletedDialogX = Range("CompletedDialog_X").Value
    Call LE_GetUserValues(CallingSheet)
End Sub
Function IE_Find_Or_Open_TEMPO() As Object
' Finds existing TEMPO window/tab in Internet Explorer
'  or, if not found, opens a new one
'
    Dim objIE As Object
    Dim shellWins As New ShellWindows
    Dim IE_TabURL As String
    
    'First, check each open window/tab for an active TEMPO session
    Set objIE = IE_Find_TEMPO()
    
    'Did we find an active TEMPO session?
    If objIE Is Nothing Then
        'No, so open a new TEMPO session
        Set objIE = CreateObject("InternetExplorer.Application")
        objIE.Visible = True
        objIE.Navigate URL_TEMPO
        Call IE_Wait(DoubleDelay)
    End If
    
    Set IE_Find_Or_Open_TEMPO = objIE
    
End Function
Function IE_Find_TEMPO() As Object
' Finds existing TEMPO window/tab in Internet Explorer
'
    Dim objIE As Object
    Dim shellWins As New ShellWindows
    Dim IE_TabURL As String
    Dim foundTEMPO As Boolean
    
    'Debug.Print "Function IE_Find_TEMPO"
    
    'First, check each open window/tab for an active TEMPO session
    foundTEMPO = False
    For Each objIE In shellWins
    
        IE_TabURL = objIE.LocationURL
        
        'Debug.Print objIE.hwnd, IE_TabURL
        
        If (IE_TabURL = URL_TEMPO) Then
            'Found a valid TEMPO URL
            foundTEMPO = True
            Exit For
        End If
        
        If (IE_TabURL = URL_LoggedOff) Then
            'Found the TEMPO "You Have Been Logged Off" URL
            foundTEMPO = True
            Exit For
        End If
        
        If Len(IE_TabURL) > Len(URL_TEMPO) Then
            If Left(IE_TabURL, Len(URL_TEMPO)) = URL_TEMPO Then
                'Found a valid TEMPO URL
                foundTEMPO = True
                Exit For
            End If
        End If
            
    Next objIE
    
    'Did we find an active TEMPO session?
    If foundTEMPO Then
        'Yes, so return the session object
        Set IE_Find_TEMPO = objIE
    Else
        'No, return the object as Nothing
        Set IE_Find_TEMPO = Nothing
    End If
    
    ' Added the delay to allow TEMPO to bring up the week data before attempting data entry.
    ' V3.10 of Timesheet Helper
    Call IE_Wait(SingleDelay)
    
End Function
Sub IE_DeleteRows_TEMPO(objIE As Object, rowIndex As Integer)
' Deletes rows from rowIndex and beyond
' Note: Infinite loop can occur if deleting the only row in TEMPO.
'       TEMPO automatically creates a new blank row, so we keep trying
'       to delete it.  Therefore, we check below whether we are deleting
'       the only row, and if so, we do not check and try to delete it again.
'       This avoids the infinite loop.
'
    Dim objElement As Object
    Dim i As Integer
    Dim altRowIndex As Integer
    Dim StartOver As Boolean

    'If we are deleting rows starting with the first row (row index 0), then
    ' start with row index 1 until there are no other rows left, then delete row index 0
    If rowIndex = 0 Then
        altRowIndex = 1
    Else
        altRowIndex = rowIndex
    End If
    'Delete the rows
    Do
        i = 0
        StartOver = False
        'Debug.Print objIE.Document.Count
        For Each objElement In objIE.Document.all
            'Debug.Print objElement.tagName, objElement.ID
            If (objElement.tagName = "SPAN") Then
                If (objElement.Title = "Delete Line") Then
                    If i = altRowIndex Then
                        objElement.Click
                        Call IE_Wait(DoubleDelay)
                        If altRowIndex > 0 Then    'avoid infinite loop
                            StartOver = True
                        End If
                        Exit For
                    Else
                        i = i + 1
                    End If
                ElseIf (objElement.Title = "Add Line") Then
                    Exit For
                End If
            End If
        Next
        'Check for special case where we still need to delete row index 0
        If (Not (StartOver)) And (altRowIndex = 1) And (rowIndex = 0) Then
            'Loop one more time to delete row index 0
            altRowIndex = 0
            StartOver = True
        End If
    Loop Until (Not (StartOver))
    
End Sub
Sub IE_EnterChargeObj_TEMPO(objIE As Object, rowIndex As Integer, theWPM As String, theChargeObj As String, theExt As String, theShift As String, theHours() As String)
' Enters Charge Object theValue in row rowIndex
'
' Rev 3.21 - Update to include the WPM Field by looking at the TD cell#s
' cell0 - Delete/Add Line
' cell1 - Add Favorite
' cell2 - ??
' cell3 - Charge Object
' cell4 - Ext
' cell5 - WPM (may not be present on a line by line basis)
' cell6-8 - ??
' cell9 - Shift
' cell10-16 - Monday - Sunday hours (Flex410)
    
    Dim objElement As Object
    Dim evt As Object
    Dim i, j As Integer
    Dim state As Integer
    Dim result As Integer
    Dim StartOver As Boolean
    Dim CellNum As Integer
    
    CellNum = -1

    Call Check_WD_Job(objIE)    'check for WD Job field (sets Found_WD_Job to "Y" or "N")
    If theShift = "" Then
        theShift = "1"
    End If
    Do
        i = 0
        state = 0
        StartOver = False
        'Debug.Print objIE.Document.Count
        For Each objElement In objIE.Document.all
            'Debug.Print objElement.tagName, objElement.ID
            'Debug.Print "ID: ", objElement.ID, "Length: ", Len(objElement.ID), "InStr: ", InStr(objElement.ID, "cell")
            If (InStr(objElement.ID, "cell") > 0) And _
               (Len(objElement.ID) > 0) Then
                CellNum = Right(objElement.ID, Len(objElement.ID) - InStr(objElement.ID, "cell") - 3)
            End If
            'Debug.Print objElement.tagName, objElement.ID, objElement.Title, "__Cell #", CellNum
            If state = 0 Then   'look for span with title "Delete Line" or "Add Line"
                If (objElement.tagName = "SPAN") Then
                    If (objElement.Title = "Delete Line") Then
                        If i = rowIndex Then
                            state = 1
                        Else
                            i = i + 1
                        End If
                    ElseIf (objElement.Title = "Add Line") Then
                        objElement.Click
                        Call IE_Wait(DoubleDelay)
                        StartOver = True
                        Exit For
                    End If
                End If
            ElseIf state = 1 Then   'look for input field for the charge object: tagName "INPUT" with role "textbox"
                If (objElement.tagName = "INPUT") Then
                    If (objElement.role = "textbox") Or _
                       (objElement.Type = "text") Then
                        state = 2
                        objElement.Focus
                        If Not (objElement.Value = UCase(theChargeObj)) Then
                            objElement.Value = theChargeObj
                            'objElement.Click
                            Set evt = objIE.Document.createEvent("HTMLEvents")
                            'Set evt = objIE.Document.createEvent("keyboardevent")
                            evt.initEvent "change", True, False
                            objElement.dispatchEvent evt
                            Call IE_SendKeys(objIE, "{TAB}")
                            Call IE_Wait(SingleDelay)
                        End If
                    End If
                End If
            ElseIf state = 2 Then   'next input field is Ext
                If (objElement.tagName = "INPUT") Then
                    If (objElement.role = "textbox") Or _
                       (objElement.Type = "text") Then
                        state = 3 ' Assume WPM Field is present
                        objElement.Focus
                        If Not (objElement.Value = UCase(theExt)) Then
                            objElement.Value = theExt
                            'objElement.Click
                            Set evt = objIE.Document.createEvent("HTMLEvents")
                            'Set evt = objIE.Document.createEvent("keyboardevent")
                            evt.initEvent "change", True, False
                            objElement.dispatchEvent evt
                            Call IE_SendKeys(objIE, "{TAB}")
                            Call IE_Wait(SingleDelay)
                        End If
                    End If
                End If
            ElseIf state = 3 Then   'next input field is WPM
                If (CellNum = 9) Then  ' Skip if no WPM
                    state = 4
                End If
                If (objElement.tagName = "INPUT") Then
                    If (objElement.role = "textbox") Or _
                       (objElement.Type = "text") Then
                        state = 4
                        objElement.Focus
                        ' We don't have a field for WD Job
                        '  set WD Job to blank, let user enter the correct code later if needed
                        If Not (objElement.Value = UCase(theWPM)) Then
                            objElement.Value = theWPM
                            'objElement.Click
                            Set evt = objIE.Document.createEvent("HTMLEvents")
                            'Set evt = objIE.Document.createEvent("keyboardevent")
                            evt.initEvent "change", True, False
                            objElement.dispatchEvent evt
                            Call IE_SendKeys(objIE, "{TAB}")
                            Call IE_Wait(SingleDelay)
                        End If
                    End If
                End If
            ElseIf state = 4 Then   'next input field is Shift
                If (objElement.tagName = "INPUT") Then
                    If (objElement.role = "textbox") Or _
                       (objElement.Type = "text") Then
                        state = 5
                        j = 0
                        objElement.Focus
                        If Not (objElement.Value = UCase(theShift)) Then
                            objElement.Value = theShift
                            'objElement.Click
                            Set evt = objIE.Document.createEvent("HTMLEvents")
                            'Set evt = objIE.Document.createEvent("keyboardevent")
                            evt.initEvent "change", True, False
                            objElement.dispatchEvent evt
                            Call IE_SendKeys(objIE, "{TAB}")
                            Call IE_Wait(SingleDelay)
                        End If
                    End If
                End If
            ElseIf state = 5 Then   'next input fields are Hours
                If (objElement.tagName = "INPUT") Then
                    If (objElement.role = "textbox") Or _
                       (objElement.Type = "text") Then
                        objElement.Focus
                        'TEMPO allows tenths of hours - compare both values as numbers rounded to 1 decimal place
                        If Not (Round(Val(objElement.Value), 1) = Round(Val(theHours(j)), 1)) Then
                            objElement.Value = theHours(j)
                            'objElement.Click
                            Set evt = objIE.Document.createEvent("HTMLEvents")
                            'Set evt = objIE.Document.createEvent("keyboardevent")
                            evt.initEvent "change", True, False
                            objElement.dispatchEvent evt
                            Call IE_SendKeys(objIE, "{TAB}")
                            Call IE_Wait(NoDelay)
                        End If
                        j = j + 1
                        If j > UBound(theHours) Then
                            state = 6
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    Loop Until (Not (StartOver))
    
    If state <> 6 Then
        Excel_Activate
        result = MsgBox("Unable to enter Charge Object for row " & rowIndex & " in TEMPO.", vbExclamation)
        IE_Finish
    End If
End Sub
Function IE_VerifyChargeObj_TEMPO(objIE As Object, rowIndex As Integer, theWPM As String, theChargeObj As String, theExt As String, theShift As String, theHours() As String) As Boolean
' Verifies the entries for the Charge Object theValue in row rowIndex
'
' Rev 3.21 - Update to include the WPM Field by looking at the TD cell#s
' cell0 - Delete/Add Line
' cell1 - Add Favorite
' cell2 - ??
' cell3 - Charge Object
' cell4 - Ext
' cell5 - WPM (may not be present on a line by line basis)
' cell6-8 - ??
' cell9 - Shift
' cell10-16 - Monday - Sunday hours (Flex410)
    
    Dim objElement As Object
    Dim evt As Object
    Dim i, j As Integer
    Dim state As Integer
    Dim result As Integer
    Dim matches As Boolean
    Dim CellNum As Integer
    
    CellNum = -1

    Call Check_WD_Job(objIE)    'check for WD Job field (sets Found_WD_Job to "Y" or "N")
    If theShift = "" Then
        theShift = "1"
    End If
    i = 0
    state = 0
    matches = True
    'Debug.Print objIE.Document.Count
    For Each objElement In objIE.Document.all
        'Debug.Print objElement.tagName, objElement.ID
        'Debug.Print "ID: ", objElement.ID, "Length: ", Len(objElement.ID), "InStr: ", InStr(objElement.ID, "cell")
        If (InStr(objElement.ID, "cell") > 0) And _
           (Len(objElement.ID) > 0) Then
                CellNum = Right(objElement.ID, Len(objElement.ID) - InStr(objElement.ID, "cell") - 3)
        End If
       'Debug.Print objElement.tagName, objElement.ID, objElement.Title, "__Cell #", CellNum
        If state = 0 Then   'look for span with title "Delete Line" or "Add Line"
            If (objElement.tagName = "SPAN") Then
                If (objElement.Title = "Delete Line") Then
                    If i = rowIndex Then
                        state = 1
                    Else
                        i = i + 1
                    End If
                End If
            End If
        ElseIf state = 1 Then   'look for input field for the charge object: tagName "INPUT" with role "textbox"
            If (objElement.tagName = "INPUT") Then
                If (objElement.role = "textbox") Or _
                   (objElement.Type = "text") Then
                    'Debug.Print state, objElement.Value, UCase(theChargeObj)
                    state = 2
                    If Not (objElement.Value = UCase(theChargeObj)) Then
                        matches = False
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 2 Then   'next input field is Ext
            If (objElement.tagName = "INPUT") Then
                If (objElement.role = "textbox") Or _
                   (objElement.Type = "text") Then
                    'Debug.Print state, objElement.Value, UCase(theExt)
                    state = 3 ' Assume WPM Field is present
                    If Not (objElement.Value = UCase(theExt)) Then
                        matches = False
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 3 Then   'next input field is WPM
            If (CellNum = 9) Then  ' Skip if no WPM
                state = 4
            End If
            If (objElement.tagName = "INPUT") Then
                If (objElement.role = "textbox") Or _
                   (objElement.Type = "text") Then
                    state = 4
                    If Not (objElement.Value = UCase(theWPM)) Then
                        matches = False
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 4 Then   'next input field is Shift
            If (objElement.tagName = "INPUT") Then
                If (objElement.role = "textbox") Or _
                   (objElement.Type = "text") Then
                    'Debug.Print state, objElement.Value, UCase(theShift)
                     state = 5
                    j = 0
                    If Not (objElement.Value = UCase(theShift)) Then
                        matches = False
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 5 Then   'next input fields are Hours
            If (objElement.tagName = "INPUT") Then
                If (objElement.role = "textbox") Or _
                   (objElement.Type = "text") Then
                    'TEMPO allows tenths of hours - compare both values as numbers rounded to 1 decimal place
                    'Debug.Print state, Round(Val(objElement.Value), 1), Round(Val(theHours(j)), 1)
                    If Not (Round(Val(objElement.Value), 1) = Round(Val(theHours(j)), 1)) Then
                        matches = False
                    End If
                    j = j + 1
                    If j > UBound(theHours) Then
                        state = 6
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    'Debug.Print "matches", rowIndex, matches
    
    IE_VerifyChargeObj_TEMPO = matches
End Function
Private Sub testIEECO()
' TEMPO must already be open at Time Entry page in an Internet Explorer window
    Dim objIE As Object
    Dim theHours(7) As String
    
    theHours(0) = "1.0"
    theHours(1) = "1.1"
    theHours(2) = "1.2"
    theHours(3) = "1.3"
    theHours(4) = "1.4"
    theHours(5) = "1.5"
    theHours(6) = "1.6"
    theHours(7) = "1.7"
    'Debug.Print LBound(theHours), UBound(theHours)
    'Get user values so URL is defined
    Call IE_GetUserValues(Instructions_ShName)
    'Get the current web browser session
    Set objIE = IE_Find_TEMPO()
    'Bring it to the front
    Call IE_Activate(objIE)
    'Test the subroutine
    Call IE_EnterChargeObj_TEMPO(IE_Find_Or_Open_TEMPO(), 0, "Test0", "Ex0", "0", theHours)
    Call IE_EnterChargeObj_TEMPO(IE_Find_Or_Open_TEMPO(), 1, "Test1", "Ex1", "1", theHours)
    Call IE_EnterChargeObj_TEMPO(IE_Find_Or_Open_TEMPO(), 2, "Test2", "Ex2", "2", theHours)
End Sub
Sub Check_WD_Job(objIE As Object)
' Checks for the "WD Job" field in TEMPO labor entry screen
'
    Dim objElement As Object
    Dim found As Boolean
    
    If (Found_WD_Job = "Y") Or (Found_WD_Job = "N") Then
        'already checked and set variable - nothing more to do
    Else
        found = False
        For Each objElement In objIE.Document.all
            If (objElement.tagName = "LABEL") Then
                If (UCase(Trim(objElement.innerText)) = "WD JOB") Then
                    found = True
                End If
            End If
        Next
        If found Then
            Found_WD_Job = "Y"
        Else
            Found_WD_Job = "N"
        End If
    End If
End Sub
Sub testCWDJ()
' TEMPO must already be open at Time Entry page in an Internet Explorer window
    Dim objIE As Object
    
    'Force subroutine to check again! (comment out to test that it "remembers")
    Found_WD_Job = ""
    'Get user values so URL is defined
    Call IE_GetUserValues(Instructions_ShName)
    'Get the current web browser session
    Set objIE = IE_Find_TEMPO()
    'Bring it to the front
    Call IE_Activate(objIE)
    'Test the subroutine
    Call Check_WD_Job(objIE)
    'Debug.Print Found_WD_Job
End Sub
Sub IE_EnterDaysOff_TEMPO(objIE As Object, dayName As String, theDate As String, dayOffValue As String)
' Enters the day off at position dayIndex (0 to 7)
'
    Dim objElement As Object
    Dim dayNumStr As String
    Dim state As Integer
    Dim result As Integer

    dayNumStr = Format(theDate, "d")
    state = 0
    For Each objElement In objIE.Document.all
        'Debug.Print objElement.tagName, objElement.ID
        If state = 0 Then   'look for label with the day name: tagName "LABEL" with innerText "Fri", for example
            If (objElement.tagName = "LABEL") Or _
               (objElement.tagName = "BDI") Then
                If (UCase(Trim(objElement.innerText)) = UCase(dayName)) Then
                    state = 1
                End If
            End If
        ElseIf state = 1 Then   'next label element must have correct day number
            If (objElement.tagName = "LABEL") Or _
               (objElement.tagName = "BDI") Then
                If (Trim(objElement.innerText) = dayNumStr) Then
                    state = 2
                Else
                    state = 0
                End If
            End If
        ElseIf state = 2 Then   'find next button: tagName "BUTTON"
            If (objElement.tagName = "BUTTON") Then
                state = 3
                If (objElement.textContent = "") Or _
                   (UCase(Trim(objElement.textContent)) = UCase(dayName)) Then
                    'button is On - textContent is empty or contains the day name (as a tooltip)
                    If dayOffValue = "" Then 'desired value is On
                        'already On, nothing to do
                    Else 'desired value is Off
                        objElement.Click 'click button to change to Off
                    End If
                Else 'button is Off (textContent = "OFF" or "FLEX")
                    If dayOffValue = "" Then 'desired value is On
                        objElement.Click 'click button to change to On
                    Else 'desired value is Off
                        'already Off, nothing to do
                    End If
                End If
                Exit For
            End If
        End If
    Next

    If state <> 3 Then
        Excel_Activate
        result = MsgBox("Unable to find day off entry for " & dayName & " " & theDate & " in TEMPO.", vbExclamation)
        IE_Finish
    End If
End Sub
Function IE_VerifyDaysOff_TEMPO(objIE As Object, dayName As String, theDate As String, dayOffValue As String) As Boolean
' Verfies the entered day off at position dayIndex (0 to 7)
'
    Dim objElement As Object
    Dim dayNumStr As String
    Dim state As Integer
    Dim result As Integer
    Dim matches As Boolean

    dayNumStr = Format(theDate, "d")
    state = 0
    matches = False
    For Each objElement In objIE.Document.all
        'Debug.Print objElement.tagName, objElement.ID
        If state = 0 Then   'look for label with the day name: tagName "LABEL" with innerText "Fri", for example
            If (objElement.tagName = "LABEL") Or _
               (objElement.tagName = "BDI") Then
                If (UCase(Trim(objElement.innerText)) = UCase(dayName)) Then
                    state = 1
                End If
            End If
        ElseIf state = 1 Then   'next label element must have correct day number
            If (objElement.tagName = "LABEL") Or _
               (objElement.tagName = "BDI") Then
                If (Trim(objElement.innerText) = dayNumStr) Then
                    state = 2
                Else
                    state = 0
                End If
            End If
        ElseIf state = 2 Then   'find next button: tagName "BUTTON"
            If (objElement.tagName = "BUTTON") Then
                state = 3
                If (objElement.textContent = "") Or _
                   (UCase(Trim(objElement.textContent)) = UCase(dayName)) Then
                    'button is On - textContent is empty or contains the day name (as a tooltip)
                    If dayOffValue = "" Then 'desired value is On
                        matches = True
                    End If
                Else 'button is Off (textContent = "OFF")
                    If dayOffValue = "" Then 'desired value is On
                    Else 'desired value is Off
                        matches = True
                    End If
                End If
                Exit For
            End If
        End If
    Next

    If state <> 3 Then
        Excel_Activate
        result = MsgBox("Unable to find day off entry for " & dayName & " " & theDate & " in TEMPO.", vbExclamation)
        IE_Finish
    End If
    
    IE_VerifyDaysOff_TEMPO = matches
End Function
Private Sub testIEEDO()
' TEMPO must already be open at Time Entry page in an Internet Explorer window
    Dim objIE As Object
    
    'Get user values so URL is defined
    Call IE_GetUserValues(Instructions_ShName)
    'Get the current web browser session
    Set objIE = IE_Find_TEMPO()
    'Bring it to the front
    Call IE_Activate(objIE)
    'Test the subroutine
    Call IE_EnterDaysOff_TEMPO(objIE, "FRI", "9/16/2016", "")
End Sub
Function IE_GetWEDate_TEMPO(objIE As Object) As String
' Returns the current Week Ending date from TEMPO in format mm/dd/yyyy
' Returns an empty string if the current Week Ending date is not found
'
    Dim objElement As Object
    Dim theStr As String
    Dim state As Integer
    Dim result As Integer
    
    theStr = ""
    state = 0
    'Look through all elements
    For Each objElement In objIE.Document.all
        'Debug.Print objElement.tagName, objElement.ID
        If state = 0 Then   'look for Payroll W/E label: tagName "LABEL" with innerText "Payroll W/E:"
            If (objElement.tagName = "LABEL") Or _
               (objElement.tagName = "BDI") Then
                If (Trim(objElement.innerText) = "Payroll W/E:") Then
                    state = 1
                End If
            End If
        ElseIf state = 1 Then   'grab next input field to get date: tagName "INPUT"
            If (objElement.tagName = "INPUT") Then
                state = 2
                theStr = Trim(objElement.Value)
                Exit For
            End If
        End If
    Next

    IE_GetWEDate_TEMPO = theStr
End Function
Function IE_GetTimeEntry_TEMPO(objIE As Object) As String
' Returns the Time Entry URL suffix from the TEMPO main page
'
    Dim objElement As Object
    Dim theStr As String
    
    theStr = ""
    'Look through all elements
    For Each objElement In objIE.Document.all
        'Debug.Print objElement.tagName, objElement.ID
        'look for Time Entry label: tagName "DIV" with title "Time Entry"
        If (objElement.tagName = "DIV") Then
            If (Trim(objElement.Title) = "Time Entry") Then
                theStr = objElement.Data - targeturl
            End If
        End If
    Next

    IE_GetTimeEntry_TEMPO = theStr
End Function
Sub testIEGTE()
    MsgBox IE_GetTimeEntry_TEMPO(IE_Find_Or_Open_TEMPO())
End Sub
Sub IE_Save_TEMPO(objIE As Object)
' Clicks the Save button after entering labor in TEMPO
'
    Dim objElement As Object
    Dim evt As Object
    Dim state As Integer
    Dim result As Integer
    
    state = 0
    'Look through all elements
    For Each objElement In objIE.Document.all
        'Debug.Print objElement.tagName, objElement.ID
        If state = 0 Then   'look for Save button: tagName "BUTTON" with innerText "Save"
            If (objElement.tagName = "BUTTON") Then
                'Debug.Print objElement.ID, Right(objElement.innerText, 4)
                If (Right(objElement.innerText, 4) = "Save") Then
                    state = 1
                    objElement.Focus
                    'objElement.Click   'this doesn't work
                    'Call IE_SendKeys(objIE, "{ENTER}")  'this doesn't work either
                    ''this doesn't work either:
                    'Set evt = objIE.Document.createEvent("HTMLEvents")
                    'evt.initEvent "onmouseover", True, False
                    'objElement.dispatchEvent evt
                    'Call IE_Wait(DoubleDelay)
                    'Set evt = objIE.Document.createEvent("HTMLEvents")
                    'evt.initEvent "onmousedown", True, False
                    'objElement.dispatchEvent evt
                    'Call IE_Wait(DoubleDelay)
                    'Set evt = objIE.Document.createEvent("HTMLEvents")
                    'evt.initEvent "onmouseup", True, False
                    'objElement.dispatchEvent evt
                    Call IE_Wait(DoubleDelay)
                    Exit For
                End If
            End If
        End If
    Next

    If state <> 1 Then
        Excel_Activate
        result = MsgBox("Unable to find the Save button in TEMPO.", vbExclamation)
        IE_Finish
    End If
End Sub
Sub IE_TimeEntry_TEMPO()
' Go to the Time Entry screen in TEMPO
'
    Dim objIE As Object
    Dim waitTime As Integer
    Dim result As Integer
    
    'Get the current web browser session
    Set objIE = IE_Find_TEMPO()
    'Bring it to the front
    If Not IE_Activate(objIE) Then
        Excel_Activate
        result = MsgBox("Unable to find the Internet Explorer window and tab containing the TEMPO Time Sheet.", vbExclamation)
        IE_Finish
    End If
    'Navigate to Time Entry
    objIE.Navigate URL_TEMPO & Suffix_Time_Entry
    'Wait for page to load
    Call IE_Wait_Until_Done(objIE)
    'Check whether page is loaded
    waitTime = 0
    Do While (IE_GetWEDate_TEMPO(objIE) = "") And (waitTime < TimeOut)
        'Wait a little longer
        Call IE_Wait(DoubleDelay)
        waitTime = waitTime + DoubleDelay
    Loop
End Sub
Sub IE_Wait(theDelaySeconds As Integer)
    If theDelaySeconds >= 0 Then
        Application.Wait DateAdd("s", theDelaySeconds, Now)
    End If
End Sub
Sub IE_Wait_Until_Done(objIE As Object)
    Dim IEbusy
    Dim IEreadyState
    Dim theType As String
    Dim waitComplete As Boolean
    
    'Debug.Print "IE_Wait_Until_Done"
    waitComplete = False
    Do
        'set default values before checking browser object
        IEbusy = True
        IEreadyState = 0
        'make sure we still have a web browser object
        theType = TypeName(objIE)
        If theType = "IWebBrowser2" Then
            IEbusy = objIE.busy
            IEreadyState = objIE.ReadyState
            'Debug.Print IEbusy, IEreadyState
        Else 'not a web browser object
            Set objIE = IE_Find_TEMPO() 're-link to IE web browser
        End If
        'check to see whether web browser is done loading the page
        If (Not IEbusy) And (IEreadyState = READYSTATE_COMPLETE) Then
            waitComplete = True
        Else
            ''dismiss dialog box in IE, if there is one
            'theType = TypeName(objIE)
            'If theType = "IWebBrowser2" Then
            '    Call IE_SendKeys(objIE, "{ENTER}")
            'End If
            'now let system do events and wait
            DoEvents
            Call IE_Wait(SingleDelay)
        End If
    Loop Until waitComplete

End Sub
Function IE_WhatsRunning() As String
' Checks what TEMPO screen is showing in Internet Explorer, returns a string identifying the name
    Dim objIE As Object
    Dim IE_TabURL As String
    Dim objElement As Object
    Dim theStr As String
'    Dim foundUserName As Boolean
    
    Set objIE = IE_Find_TEMPO()
    Call IE_Wait_Until_Done(objIE)
    IE_TabURL = objIE.LocationURL
'    foundUserName = False
    If (IE_TabURL = URL_TEMPO) Then
        'check for login screen
        If objIE.LocationName = "Logon" Then
            IE_WhatsRunning = "TEMPO Login Page"
        Else
            IE_WhatsRunning = "TEMPO Welcome Page"
        End If
        Exit Function
    End If
    If (IE_TabURL = URL_TEMPO & Suffix_Shell_Home) Then
        IE_WhatsRunning = "TEMPO Welcome Page"
        Exit Function
    End If
    If (IE_TabURL = URL_TEMPO & Suffix_Time_Entry) Then
        IE_WhatsRunning = "TEMPO Time Entry Page"
        Exit Function
    End If
    If (IE_TabURL = URL_LoggedOff) Then
        IE_WhatsRunning = "TEMPO Logged Off Page"
        Exit Function
    End If
    If Len(IE_TabURL) > Len(URL_TEMPO) Then
        If Left(IE_TabURL, Len(URL_TEMPO)) = URL_TEMPO Then
            IE_WhatsRunning = "TEMPO Other Page"
            Exit Function
        End If
    End If
    IE_WhatsRunning = "NONE"
End Function
Function IE_Activate(objIE As Object) As Boolean
'
'Bring Internet Explorer window to the front and select the desired tab
'
    Dim windowName
    Dim windowNameLength
    Dim winTitleBuf As String * 255
    Dim retLong As Long
    Dim foundTab As Boolean
    Dim loopCount As Integer
    Const loopLimit As Integer = 256
    
    'Debug.Print "Function IE_Activate"
    
    foundTab = False
    objIE.Visible = True
    windowName = objIE.Document.Title
    windowNameLength = Len(windowName)
    'this doesn't work when TEMPO is in a tab in IE that is not currently selected
    'but we'll wrap it with an error handler so we can continue anyway
    On Error Resume Next
    If windowNameLength > 0 Then
        AppActivate (windowName)
    End If
    AppActivate objIE
    On Error GoTo 0
    'make sure window is topmost
    BringWindowToTop objIE.hwnd
    'select desired tab if multiple tabs are present
    If windowNameLength > 0 Then
        loopCount = 0
        Do
            'get window title of IE window (includes " - Internet Explorer" at end)
            retLong = GetWindowText(objIE.hwnd, winTitleBuf, 255)
            'Debug.Print retLong, winTitleBuf
            If retLong >= windowNameLength Then
                'check for title match to the length of the desired window name
                If Left(winTitleBuf, windowNameLength) = windowName Then
                    'it is a match!
                    foundTab = True
                End If
            End If
            If Not foundTab Then
                'haven't found the correct tab yet
                'Send Ctrl+Tab to IE to switch to next tab
                Call IE_SendKeys(objIE, "^{TAB}")
                'Wait one second so window has time to refresh
                Call IE_Wait(1)
            End If
            'keep track of number of times through loop to prevent infinite loop
            loopCount = loopCount + 1
        Loop Until foundTab Or (loopCount > loopLimit)
    End If
    'return boolean indicating whether correct tab was found
    IE_Activate = foundTab
End Function
Sub IE_SendKeys(objIE As Object, theString As String)
'
'Bring Internet Explorer window to the front
'
    Dim windowName
    
    objIE.Visible = True
    windowName = objIE.Document.Title
    'this doesn't work when TEMPO is in a tab in IE that is not currently selected
    'but we'll wrap it with an error handler so we can continue anyway
    On Error Resume Next
    AppActivate (windowName)
    AppActivate objIE
    On Error GoTo 0
    'make sure the window is still in front
    BringWindowToTop objIE.hwnd
    'and send the keys
    'workaround for SendKeys bug (see https://support.microsoft.com/en-us/kb/179987)
    DoEvents
    Get_Keyboard_States
    Application.SendKeys (theString), True
    Set_Keyboard_States
    DoEvents
End Sub
Sub IE_Finish()
'
' Quits
'
    End
End Sub

