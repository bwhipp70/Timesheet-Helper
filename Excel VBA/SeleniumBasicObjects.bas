'Macro Module: SeleniumBasicObjects
'Last Updated: 2021-10-12 WRH

'This macro module is used in UpTEMPO

'Purpose: Interface with Edge and Chrome to enter labor in TEMPO

'Recent changes (in reverse chronological order):
' 2021-10-12 WRH: Created, based on InternetExplorerObjects
'
' Thanks to:
'
'   https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement
'   https://excelmacromastery.com/vba-error-handling/
'
Option Explicit
'
' To get Selenium Type Library: go to the Tools menu at the top of the VBA
'  editor, click on References in the menu, and scroll down to find
'  the “Selenium Type Library” reference. Click the checkbox to the
'  left of it, and then click OK.
'
' Drop-down entry for Internet Explorer (VBA)
'
Public Const SB_Edge_BrowserDriver = "Edge (SeleniumBasic)"
Public Const SB_Chrome_BrowserDriver = "Chrome (SeleniumBasic)"
'
' Keep track of whether TEMPO time entry sheet has "WD Job" field and "TVL" field
'  default is empty string (need to check), "Y" if yes, and "N" if no
'
Dim Found_WD_Job As String
Dim Found_WPM As String
Dim Found_TVL As String
'
' Private global variables
'
Private SB_numElements As Long
Private SB_elementIndex As Long
Private SB_rowCount As Integer
Private SB_mismatch As String
'
Sub SB_EnterLabor(CallingSheet)
' Logs in to TEMPO and enters labor from Labor Worksheet
'
    Dim driver As New WebDriver
    Dim result As Integer
    Dim FirstLaborRow, LastLaborRow
    Dim colTotalHours
    Dim i As Integer
    Dim TEMPO_WEdate, WEdate As Date
    Dim iRow
    Dim iEntries As Integer
    Dim LastEntryRow
    Dim theHours(7) As String
    Dim entriesMatch As Boolean
    Dim loopCount As Integer
    Const loopLimit As Integer = 3

' Add variables for the Status Bar
    Dim sBar_Name1 As String
    Dim sBar_Name2 As String
    Dim FLR As Integer
    Dim LLR As Integer
    Dim TotalRows As Integer
    Dim TotalElements As Integer
    Dim MajorElements As Integer
    Dim MinorElements As Integer
    Dim sBar As Boolean
    
    sBar_Name1 = "TEMPO Uploading"
    sBar_Name2 = "Initialize"
    
' Find out how many rows of labor
    If CallingSheet = Labor_Flex410_ShName Then
        FLR = FirstLaborRow_Flex410
        LLR = LastLaborRow_Flex410
    ElseIf CallingSheet = Labor_Flex980_ShName Then
        FLR = FirstLaborRow_Flex980
        LLR = LastLaborRow_Flex980
    ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
        FLR = FirstLaborRow_Flex980_2weeks
        LLR = LastLaborRow_Flex980_2weeks
    End If
    
    iRow = FLR
    iEntries = 0
    'Find last row with a labor entry
    LastEntryRow = LLR
    Do While (LastEntryRow >= FirstLaborRow) And _
        (Sheets(CallingSheet).Cells(LastEntryRow, 3).Value = "")
        LastEntryRow = LastEntryRow - 1
    Loop
 
    TotalRows = LastEntryRow + 1 - FLR
    
    ' Start Tempo -> (5) Open, Navigate, Check, TE Page, W/E
    ' Enter Days Off -> (7) M, T, W, Th, F, S, S
    ' Enter Charge Obj -> State 1, 2, 4, 5, 6, 7 for 7x [j=0 to 6] = 12 * Total Number of Rows
    ' Validate Days Off -> (7) M, T, W, Th, F, S, S
    ' Validate Charge Obj -> State 1, 2, 4, 5, 6, 7 for 7x [j=0 to 6] = 12 * Total Number of Rows
    
    TotalElements = 5 + 7 + (12 * TotalRows) + 7 + (12 * TotalRows)
    
    ' Format for section call:
    '   sBar_Name1 = TEMPO Init
    '   sBar_Name2 = Open TEMPO
    '   MajorElements = 0 ' Starting # of Elements for Major Grouping
    '   MinorElement = 0  ' Minor always starts at 0
    
    '   sBar = StatusBar_Draw2(sBar_Name1, Round(((MajorElements+MinorElements) / TotalElements) * 100, 0), sBar_Name2, Round((MinorElements/##) * 100, 0))
    '   where ## is the number of minor elements per major element
    
    ' End of Status Bar Initialization
    
    'Open TEMPO in selected web browser
    
    ' Status Bar
    sBar = StatusBar_Draw2("TEMPO Init", Round(((0 + 1) / TotalElements) * 100, 0), "Open Browser", Round((1 / 5) * 100, 0))
    
    On Error GoTo SBDriverError
    
    If theBrowserDriver = SB_Edge_BrowserDriver Then
        driver.Start "Edge"
    ElseIf theBrowserDriver = SB_Chrome_BrowserDriver Then
        driver.Start "Chrome"
    Else
        Call Debug_Warn_User("SB_EnterLabor", "Unknown Browser Driver")
        SB_Finish
    End If
    
    On Error GoTo 0
    
    'Navigate to TEMPO URL
    
    ' Status Bar
    sBar = StatusBar_Draw2("TEMPO Init", Round(((0 + 2) / TotalElements) * 100, 0), "Navigate", Round((2 / 5) * 100, 0))
    
    driver.Get URL_TEMPO
    
    'check for TEMPO Authentication page (from non LMI)
    ' Added in TS Helper
    If SB_WhatsRunning(driver) = "TEMPO Authentication Page" Then
        Excel_Activate
        result = MsgBox("Authenticate Then Press OK." & Chr(13) & _
                        "NOTE: The RSA Token only works once per minute.", vbExclamation)
    End If
    
    'check for TEMPO Login page
    
    ' Status Bar
    sBar = StatusBar_Draw2("TEMPO Init", Round(((0 + 3) / TotalElements) * 100, 0), "Check Login", Round((3 / 5) * 100, 0))
    
    If SB_WhatsRunning(driver) = "TEMPO Login Page" Then
        'Close this browser window
        driver.Close
        'Open TEMPO (again) in selected web browser
        driver.Start
        'Navigate to TEMPO URL
        driver.Get URL_TEMPO
    End If
    
    'go to TEMPO Time Entry page
    
    ' Status Bar
    sBar = StatusBar_Draw2("TEMPO Init", Round(((0 + 4) / TotalElements) * 100, 0), "Time Entry Page", Round((4 / 5) * 100, 0))
    
    If SB_WhatsRunning(driver) = "TEMPO Welcome Page" Then
        Call SB_TimeEntry_TEMPO(driver)
    ElseIf SB_WhatsRunning(driver) = "TEMPO Other Page" Then
        Call SB_TimeEntry_TEMPO(driver)
    End If
    
    
    ' Status Bar
    sBar = StatusBar_Draw2("TEMPO Init", Round(((0 + 5) / TotalElements) * 100, 0), "Check W/E", Round((5 / 5) * 100, 0))
    
    If SB_WhatsRunning(driver) = "TEMPO Time Entry Page" Then
        'configure first and last labor rows for the selected sheet
        If CallingSheet = Labor_Flex410_ShName Then
            FirstLaborRow = FirstLaborRow_Flex410
            LastLaborRow = LastLaborRow_Flex410
            colTotalHours = 15  ' Modified from UpTempo (was 19)
        ElseIf CallingSheet = Labor_Flex980_ShName Then
            FirstLaborRow = FirstLaborRow_Flex980
            LastLaborRow = LastLaborRow_Flex980
            colTotalHours = 15  ' Same as UpTempo
        ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
            FirstLaborRow = FirstLaborRow_Flex980_2weeks
            LastLaborRow = LastLaborRow_Flex980_2weeks
            colTotalHours = 21  ' Same as UpTempo
        Else
            Call Debug_Warn_User("SB_EnterLabor", "Unknown CallingSheet")
        End If
        'check whether week ending date in TEMPO matches this week's date
        If (CallingSheet = Labor_Flex410_ShName) Or _
            (CallingSheet = Labor_Flex980_ShName) Or _
            (CallingSheet = Labor_Flex980_2weeks_ShName) Then
            TEMPO_WEdate = SB_GetWEDate_TEMPO(driver)
            If TEMPO_WEdate = 0 Then
                Excel_Activate
                result = MsgBox("Unable to find Payroll W/E date in TEMPO.", vbExclamation)
                SB_Finish
            End If
            If (CallingSheet = Labor_Flex410_ShName) Then
                WEdate = CDate(Sheets(CallingSheet).Range("BH10").Value) ' Modified from UpTempo
            ElseIf (CallingSheet = Labor_Flex980_ShName) Then
                WEdate = CDate(Sheets(CallingSheet).Range("BH10").Value) + 2 ' Modified from UpTempo
            ElseIf (CallingSheet = Labor_Flex980_2weeks_ShName) Then
                WEdate = CDate(Sheets(CallingSheet).Range("BH10").Value) + 2 ' Modified from UpTempo
            End If
            If TEMPO_WEdate <> WEdate Then
                Excel_Activate
                Sheets(CallingSheet).Select
                If (CallingSheet = Labor_Flex410_ShName) Then
                    Range("BH10").Select ' Modified from UpTempo
                ElseIf (CallingSheet = Labor_Flex980_ShName) Then
                    Range("BH10").Select ' Modified from UpTempo
                ElseIf (CallingSheet = Labor_Flex980_2weeks_ShName) Then
                    Range("BH10").Select ' Modified from UpTempo
                End If
                result = MsgBox("The week ending (W/E) date in TEMPO (" & TEMPO_WEdate & _
                        ") does not match the Current Week Ending date." & Chr(13) & _
                        "Please verify the week ending dates and try again.", vbExclamation)
                ' SB_Finish
            End If
        Else
            Call Debug_Warn_User("SB_EnterLabor", "Unknown CallingSheet")
        End If
'
'
' Try to get the date again..  Giving the user the chance to change the W/E date.
        If (CallingSheet = Labor_Flex410_ShName) Or _
            (CallingSheet = Labor_Flex980_ShName) Or _
            (CallingSheet = Labor_Flex980_2weeks_ShName) Then
            TEMPO_WEdate = SB_GetWEDate_TEMPO(driver)
            If TEMPO_WEdate = 0 Then
                Excel_Activate
                result = MsgBox("Unable to find Payroll W/E date in TEMPO.", vbExclamation)
                SB_Finish
            End If
            If (CallingSheet = Labor_Flex410_ShName) Then
                WEdate = CDate(Sheets(CallingSheet).Range("BH10").Value) ' Modified from UpTempo
            ElseIf (CallingSheet = Labor_Flex980_ShName) Then
                WEdate = CDate(Sheets(CallingSheet).Range("BH10").Value) + 2 ' Modified from UpTempo
            ElseIf (CallingSheet = Labor_Flex980_2weeks_ShName) Then
                WEdate = CDate(Sheets(CallingSheet).Range("BH10").Value) + 2 ' Modified from UpTempo
            End If
            If TEMPO_WEdate <> WEdate Then
                Excel_Activate
                Sheets(CallingSheet).Select
                If (CallingSheet = Labor_Flex410_ShName) Then
                    Range("BH10").Select ' Modified from UpTempo
                ElseIf (CallingSheet = Labor_Flex980_ShName) Then
                    Range("BH10").Select ' Modified from UpTempo
                ElseIf (CallingSheet = Labor_Flex980_2weeks_ShName) Then
                    Range("BH10").Select ' Modified from UpTempo
                End If
                result = MsgBox("The week ending (W/E) date in TEMPO (" & TEMPO_WEdate & _
                        ") STILL does not match (second try) the Current Week Ending date." & Chr(13) & _
                        "Stopping Macro, please try again.", vbExclamation)
                SB_Finish
            End If
        End If
' Did the second date try work???


        'perform entries, then verify them
        loopCount = 0
        Do
            'enter days off
            If CallingSheet = Labor_Flex410_ShName Then
' Modified from UpTempo
                For i = 0 To 6
                    Call SB_EnterDaysOff_TEMPO(driver, Sheets(CallingSheet).Cells(4, 8 + i).Value, _
                        Sheets(CallingSheet).Cells(5, 8 + i).Value, Sheets(CallingSheet).Cells(8, 8 + i).Value)
                    ' Status Bar
                    sBar = StatusBar_Draw2("Enter Days Off", Round(((5 + (i + 1)) / TotalElements) * 100, 0), "Days", Round(((i + 1) / 7) * 100, 0))
                    MajorElements = 5 + (i + 1)
                    Next i
            ElseIf CallingSheet = Labor_Flex980_ShName Then
                For i = 0 To 7
                    Call SB_EnterDaysOff_TEMPO(driver, Sheets(CallingSheet).Cells(4, 7 + i).Value, _
                        Sheets(CallingSheet).Cells(5, 7 + i).Value, Sheets(CallingSheet).Cells(8, 7 + i).Value)
                Next i
            ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                For i = 0 To 7
                    Call SB_EnterDaysOff_TEMPO(driver, Sheets(CallingSheet).Cells(4, 13 + i).Value, _
                        Sheets(CallingSheet).Cells(5, 13 + i).Value, Sheets(CallingSheet).Cells(8, 13 + i).Value)
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
                    If CallingSheet = Labor_Flex410_ShName Then
                        theHours(0) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Mon
                        theHours(1) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Tue
                        theHours(2) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Wed
                        theHours(3) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Thu
                        theHours(4) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Fri
                        theHours(5) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Sat
                        theHours(6) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sun
                        theHours(7) = ""
                        Call SB_EnterChargeObj_TEMPO(driver, _
                            iEntries, _
                            Sheets(CallingSheet).Cells(iRow, 3).Value, _
                            Sheets(CallingSheet).Cells(iRow, 5).Value, _
                            Sheets(CallingSheet).Cells(iRow, 6).Value, _
                            theHours, _
                            6, _
                            Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                            MajorElements, loopCount, TotalElements, TotalRows, _
                            "")
                    ElseIf CallingSheet = Labor_Flex980_ShName Then
                        theHours(0) = Sheets(CallingSheet).Cells(iRow, 7).Value 'Fri
                        theHours(1) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Sat
                        theHours(2) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Sun
                        theHours(3) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Mon
                        theHours(4) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Tue
                        theHours(5) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Wed
                        theHours(6) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Thu
                        theHours(7) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Fri
'****                        Call SB_EnterChargeObj_TEMPO(driver, iEntries, Sheets(CallingSheet).Cells(iRow, 3).Value, _
                            Sheets(CallingSheet).Cells(iRow, 5).Value, Sheets(CallingSheet).Cells(iRow, 6).Value, theHours, 7)
                    ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                        theHours(0) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Fri
                        theHours(1) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sat
                        theHours(2) = Sheets(CallingSheet).Cells(iRow, 15).Value 'Sun
                        theHours(3) = Sheets(CallingSheet).Cells(iRow, 16).Value 'Mon
                        theHours(4) = Sheets(CallingSheet).Cells(iRow, 17).Value 'Tue
                        theHours(5) = Sheets(CallingSheet).Cells(iRow, 18).Value 'Wed
                        theHours(6) = Sheets(CallingSheet).Cells(iRow, 19).Value 'Thu
                        theHours(7) = Sheets(CallingSheet).Cells(iRow, 20).Value 'Fri
'****                        Call SB_EnterChargeObj_TEMPO(driver, iEntries, Sheets(CallingSheet).Cells(iRow, 3).Value, _
                            Sheets(CallingSheet).Cells(iRow, 5).Value, Sheets(CallingSheet).Cells(iRow, 6).Value, theHours, 7)
                    Else
                        Call Debug_Warn_User("SB_EnterLabor", "Unknown CallingSheet")
                    End If
                    iEntries = iEntries + 1
                End If
                iRow = iRow + 1
            Loop Until iRow > LastEntryRow
            'Status Bar
            MajorElements = 12 + (12 * TotalRows)
            'delete extra labor lines at end
            Call SB_DeleteRows_TEMPO(driver, iEntries)
            'pause briefly
            Call SB_Wait(driver, 1)
            'begin verification of entered data
            entriesMatch = True  'start as true, change to false if any entry does not match
            'verify days off
            If CallingSheet = Labor_Flex410_ShName Then
' Modified from UpTempo
                For i = 0 To 6
                    If entriesMatch Then
                        entriesMatch = SB_VerifyDaysOff_TEMPO(driver, Sheets(CallingSheet).Cells(4, 8 + i).Value, _
                            Sheets(CallingSheet).Cells(5, 8 + i).Value, Sheets(CallingSheet).Cells(8, 8 + i).Value)
                        ' Status Bar
                        sBar = StatusBar_Draw2("Checking Days Off", Round(((MajorElements + (i + 1)) / TotalElements) * 100, 0), "Days", Round(((i + 1) / 7) * 100, 0))
                    End If
                Next i
            ElseIf CallingSheet = Labor_Flex980_ShName Then
                For i = 0 To 7
                    If entriesMatch Then
                        entriesMatch = SB_VerifyDaysOff_TEMPO(driver, Sheets(CallingSheet).Cells(4, 7 + i).Value, _
                            Sheets(CallingSheet).Cells(5, 7 + i).Value, Sheets(CallingSheet).Cells(8, 7 + i).Value)
                    End If
                Next i
            ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                For i = 0 To 7
                    If entriesMatch Then
                        entriesMatch = SB_VerifyDaysOff_TEMPO(driver, Sheets(CallingSheet).Cells(4, 13 + i).Value, _
                            Sheets(CallingSheet).Cells(5, 13 + i).Value, Sheets(CallingSheet).Cells(8, 13 + i).Value)
                    End If
                Next i
            Else
                Call Debug_Warn_User("SB_EnterLabor", "Unknown CallingSheet")
            End If
            ' Status Bar
            MajorElements = MajorElements + 7
            'prepare to verify labor
            iRow = FirstLaborRow
            iEntries = 0
            Call SB_VerifyChargeObj_TEMPO_Init(driver)
            'verify labor
            Do
                If entriesMatch Then
                    'Check whether we're entering all labor
                    ' always enter labor lines with non-zero total hours
                    If (AllLaborX <> "") Or _
                        (Sheets(CallingSheet).Cells(iRow, colTotalHours).Value <> "") Then
                        If CallingSheet = Labor_Flex410_ShName Then
                            theHours(0) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Mon
                            theHours(1) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Tue
                            theHours(2) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Wed
                            theHours(3) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Thu
                            theHours(4) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Fri
                            theHours(5) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Sat
                            theHours(6) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sun
                            theHours(7) = ""
                            entriesMatch = SB_VerifyChargeObj_TEMPO(driver, _
                                iEntries, _
                                Sheets(CallingSheet).Cells(iRow, 3).Value, _
                                Sheets(CallingSheet).Cells(iRow, 5).Value, _
                                Sheets(CallingSheet).Cells(iRow, 6).Value, _
                                theHours, _
                                6, _
                                Left(Sheets(CallingSheet).Cells(iRow, 2).Value, 10), _
                                MajorElements, loopCount, TotalElements, TotalRows, _
                                "")
                        ElseIf CallingSheet = Labor_Flex980_ShName Then
                            theHours(0) = Sheets(CallingSheet).Cells(iRow, 7).Value 'Fri
                            theHours(1) = Sheets(CallingSheet).Cells(iRow, 8).Value 'Sat
                            theHours(2) = Sheets(CallingSheet).Cells(iRow, 9).Value 'Sun
                            theHours(3) = Sheets(CallingSheet).Cells(iRow, 10).Value 'Mon
                            theHours(4) = Sheets(CallingSheet).Cells(iRow, 11).Value 'Tue
                            theHours(5) = Sheets(CallingSheet).Cells(iRow, 12).Value 'Wed
                            theHours(6) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Thu
                            theHours(7) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Fri
'***                            entriesMatch = SB_VerifyChargeObj_TEMPO(driver, iEntries, Sheets(CallingSheet).Cells(iRow, 3).Value, _
                                Sheets(CallingSheet).Cells(iRow, 5).Value, Sheets(CallingSheet).Cells(iRow, 6).Value, theHours, 7)
                        ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                            theHours(0) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Fri
                            theHours(1) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sat
                            theHours(2) = Sheets(CallingSheet).Cells(iRow, 15).Value 'Sun
                            theHours(3) = Sheets(CallingSheet).Cells(iRow, 16).Value 'Mon
                            theHours(4) = Sheets(CallingSheet).Cells(iRow, 17).Value 'Tue
                            theHours(5) = Sheets(CallingSheet).Cells(iRow, 18).Value 'Wed
                            theHours(6) = Sheets(CallingSheet).Cells(iRow, 19).Value 'Thu
                            theHours(7) = Sheets(CallingSheet).Cells(iRow, 20).Value 'Fri
'***                            entriesMatch = SB_VerifyChargeObj_TEMPO(driver, iEntries, Sheets(CallingSheet).Cells(iRow, 3).Value, _
                                Sheets(CallingSheet).Cells(iRow, 5).Value, Sheets(CallingSheet).Cells(iRow, 6).Value, theHours, 7)
                        Else
                            Call Debug_Warn_User("SB_EnterLabor", "Unknown CallingSheet")
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
        
        ' Status Bar
        StatusBar_Clear
        
        'Unable to click the save button at this time, but don't want to automatically click it anyway
        ' - this allows user to review labor before TEMPO combines and rearranges it!
        'pause briefly before returning to Excel
        Call SB_Wait(driver, 1)
        Excel_Activate
        If entriesMatch Then
            'Must show this dialog!  After dialog is closed, browser will close too!
            result = MsgBox("DO NOT USE THIS TO SAVE YOUR LABOR!" & Chr(10) & Chr(10) & _
                "This is for Educational Purposes Only.", vbExclamation)
            result = MsgBox("Labor entry completed!" & Chr(10) & Chr(10) & _
                "Review the labor; if correct, click the Save button in TEMPO." & Chr(10) & Chr(10) & _
                "The TEMPO window will be closed when you click OK," & Chr(10) & _
                "and all unsaved changes will be lost!", vbExclamation)
            result = MsgBox("Last chance!!!!" & Chr(10) & Chr(10) & _
                "The TEMPO window will be closed when you click OK," & Chr(10) & _
                "and all unsaved changes will be lost!", vbExclamation)
        Else
            result = MsgBox("Unable to enter labor correctly!" & Chr(10) & Chr(10) & _
                "Tried " & loopLimit & " times and found at least one incorrect entry each time:" & Chr(10) & _
                SB_mismatch, vbExclamation)
        End If
'
'
    Else
        'not at Attendance & Labor Input page
        Excel_Activate
        result = MsgBox("Unable to get to TEMPO Time Entry page.", vbExclamation)
        SB_Finish
    End If
    
    driver.Close

Exit Sub

SBDriverError:
    If Debug_Warn Then Debug.Print "Error!  Error #" & Err.Number & ", Description: " & Err.Description
    If Err.Number = 0 Then
        MsgBox "Missing the browser specific SeleniumBasic Driver." & vbCrLf & _
                "Refer to the Educational Mode Tab for more information."
    Else
        MsgBox "Wrong version of the browser specific SeleniumBasic Driver." & vbCrLf & _
                "Refer to the Educational Mode Tab for the link to download a newer version."
    End If
    'Don't continue, end the Macro
    SB_Finish

End Sub

Sub SB_DeleteRows_TEMPO(driver As WebDriver, rowIndex As Integer)
' Deletes rows from rowIndex and beyond
' Note: Infinite loop can occur if deleting the only row in TEMPO.
'       TEMPO automatically creates a new blank row, so we keep trying
'       to delete it.  Therefore, we check below whether we are deleting
'       the only row, and if so, we do not check and try to delete it again.
'       This avoids the infinite loop.
'
    Dim objElement As WebElement
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
        For Each objElement In driver.FindElementsByXPath("//span[@title='Delete Line'] | //span[@title='Add Line']")
            If Debug_Warn Then Debug.Print objElement.tagname, objElement.Text
            If (objElement.Attribute("title") = "Delete Line") Then
                If i = altRowIndex Then
                    objElement.Click
                    Call SB_Wait(driver, DoubleDelay)
                    If altRowIndex > 0 Then    'avoid infinite loop
                        StartOver = True
                    End If
                    Exit For
                Else
                    i = i + 1
                End If
            ElseIf (objElement.Attribute("title") = "Add Line") Then
                Exit For
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

Sub SB_EnterChargeObj_TEMPO(driver As WebDriver, rowIndex As Integer, theChargeObj As String, theExt As String, theShift As String, theHours() As String, theHoursUBound As Integer, theWPM As String, MajorElements As Integer, loopCount As Integer, TotalElements As Integer, TotalRows As Integer, Optional ByVal sTVL As String = "")
' Enters Charge Object theValue in row rowIndex
'
' Rev 4.00 Cell Tab Num
' State 1 - Charge Object - 23
' State 2 - Extension - 25
' State 3 - WD Job - ??
' State 4 - WPM - 27
' State 5 - Shift - 31
' State 6 - TVL - ??
' State 7 - M - 41 to Sun - 47
' Can view the Tab Index numbers (in Edge) -> F12 -> Elements
' Needed to find a tag that would highlight the existence of the WPM field when it pops up after
' enetering a WPM enabled charge object.
'
' Shift - 31
' Rev 3.23 - TEMPO Cell# update on 6/22/2021
' cell0 - Delete/Add Line
' cell1 - Add Favorite
' cell2 - ??
' cell3 - Charge Object
' cell4 - ??
' cell5 - Ext
' cell6 - ??
' cell7 - WPM (may not be present on a line by line basis)
' cell8-10 - ??
' cell11 - Shift
' cell12-13 - ??
' cell14-20 - Monday - Sunday hours (Flex410)
    
    Dim objElement As WebElement
    Dim evt As Object
    Dim i, j As Integer
    Dim state As Integer
    Dim result As Integer
    Dim StartOver As Boolean
    Dim bTVL As Boolean
    Dim CellNum As Integer
    Dim ExtRedo As Boolean
    Dim sBar As Boolean
    
    CellNum = -1
    
    Call SB_Check_WD_Job(driver)    'check for WD Job field (sets Found_WD_Job to "Y" or "N")
    Call SB_Check_WPM(driver)       'check for WPM field (sets Found_WPM to "Y" or "N")
    Call SB_Check_TVL(driver)       'check for TVL field (sets Found_TVL to "Y" or "N")
    If theShift = "" Then
        theShift = "1"
    End If
    If sTVL = "" Then
        bTVL = False
    Else
        bTVL = True
    End If
    
' Debug statements for test
    Found_WD_Job = "N"
    Found_WPM = "N"
    Found_TVL = "N"
    ExtRedo = True
    
    driver.Window.Activate
    Do
        i = 0
        state = 0
        StartOver = False
        For Each objElement In driver.FindElementsByXPath("//span[@title='Delete Line'] | //span[@title='Add Line'] | //input | //div[contains(@class,'sapMCbMark')]")
            On Error Resume Next
                CellNum = Right(objElement.Attribute("tabindex"), 2)
            On Error GoTo 0
            If Debug_Warn Then
                Debug.Print objElement.tagname, objElement.Text, objElement.Attribute("title")
                Debug.Print "Max Length - ", objElement.Attribute("maxlength")
                Debug.Print "Tab Index - ", objElement.Attribute("tabindex")
                Debug.Print "Cell Number", CellNum
            End If
            If CellNum = 27 And state = 5 Then ' Qualify the WPM field for only the line being entered/checked
                state = 4
            End If
            If state = 0 Then   'look for span with title "Delete Line" or "Add Line"
                If (objElement.tagname = "span") Then
                    If (objElement.Attribute("title") = "Delete Line") Then
                        If i = rowIndex Then
                            state = 1
                        'Status Bar
                        sBar = StatusBar_Draw2("Time Entry", Round(((MajorElements + ((i * 12) + 1)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (i + 1) & " of " & TotalRows, Round((1 / 12) * 100, 0))
                         Else
                            i = i + 1
                        End If
                    ElseIf (objElement.Attribute("title") = "Add Line") Then
                        objElement.Click
                        Call SB_Wait(driver, DoubleDelay)
                        Call SB_Wait(driver, DoubleDelay) ' Added second delay due to occasionally missing this
                        StartOver = True
                        Exit For
                    End If
                End If
            ElseIf state = 1 Then   'look for input field for the charge object: tagName "INPUT" with role "textbox"
                If (objElement.tagname = "input") Then
                    If (objElement.Attribute("role") = "textbox") Or _
                       (objElement.Attribute("type") = "text") Then
                        state = 2
                        'Status Bar
                        sBar = StatusBar_Draw2("Time Entry", Round(((MajorElements + ((i * 12) + 2)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (i + 1) & " of " & TotalRows, Round((2 / 12) * 100, 0))
                        objElement.SendKeys ("") 'set focus to this element
                        If Not (objElement.Value = UCase(theChargeObj)) Then
                            If Not (objElement.Value = "") Then
                                objElement.Clear
                            End If
                            Call SB_Wait(driver, SingleDelay)
                            objElement.SendKeys (theChargeObj)
                            'Call SB_Wait(driver,SingleDelay)
                            'Need to start over or else we will get a stale data error
                            StartOver = True
                            Exit For
                        End If
                    End If
                End If
            ElseIf state = 2 Then   'next input field is Ext
                If (objElement.tagname = "input") Then
                    If (objElement.Attribute("role") = "textbox") Or _
                       (objElement.Attribute("type") = "text") Then
                        If Found_WD_Job = "Y" Then
                            state = 3 'need to handle WD Job field next
                        Else
                            'no WD Job field - skip it
                            If Found_WPM = "Y" Then
                                state = 4 'need to handle WPM field next
                            Else
                                state = 5 'no WPM field - skip it
                            End If
                        End If
'                        state = 4 'Assume WPM field is present
                        'Status Bar
                        sBar = StatusBar_Draw2("Time Entry", Round(((MajorElements + ((i * 12) + 3)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (i + 1) & " of " & TotalRows, Round((3 / 12) * 100, 0))
                        objElement.SendKeys ("") 'set focus to this element
                        If Not (objElement.Value = UCase(theExt)) Then
                            If Not (objElement.Value = "") Then
                                objElement.Clear
                            End If
                            Call SB_Wait(driver, SingleDelay)
                            objElement.SendKeys (theExt)
                            Call SB_Wait(driver, SingleDelay)
                            StartOver = True
                            Exit For
                        End If
                        If ExtRedo Then
                           ExtRedo = False
                           Call SB_Wait(driver, DoubleDelay) ' Need to pause to let the WPM field appear
                           StartOver = True
                           Exit For
                        End If
                    End If
                End If
            ElseIf state = 3 Then   'next input field is WD Job
                If (objElement.tagname = "input") Then
                    If (objElement.Attribute("role") = "textbox") Or _
                       (objElement.Attribute("type") = "text") Then
                        If Found_WPM = "Y" Then
                            state = 4 'need to handle WPM field next
                        Else
                            state = 5 'no WPM field - skip it
                        End If
                        objElement.SendKeys ("") 'set focus to this element
                        ' UpTEMPO doesn't have a field for WD Job
                        '  set WD Job to blank, let user enter the correct code later if needed
                        If Not (objElement.Value = "") Then
                            objElement.Clear
                            Call SB_Wait(driver, SingleDelay)
                        End If
                    End If
                End If
            ElseIf state = 4 Then   'next input field is WPM
                    state = 5
                    CellNum = -1
                    'Status Bar
                     sBar = StatusBar_Draw2("Time Entry", Round(((MajorElements + ((i * 12) + 4)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (i + 1) & " of " & TotalRows, Round((4 / 12) * 100, 0))
                    If (objElement.tagname = "input") Then
                        If (objElement.Attribute("role") = "textbox") Or _
                           (objElement.Attribute("type") = "text") Then
                            state = 5
                            objElement.SendKeys ("") 'set focus to this element
                            ' UpTEMPO doesn't have a field for WPM
                            '  set WPM to blank, let user enter the correct code later if needed
                            If Not (objElement.Value = UCase(theWPM)) Then
                                If Not (objElement.Value = "") Then
                                    objElement.Clear
                                End If
                                Call SB_Wait(driver, SingleDelay)
                                objElement.SendKeys (theWPM)
                                Call SB_Wait(driver, SingleDelay)
                            End If
                        End If
                    End If
            ElseIf state = 5 Then   'next input field is Shift
                If (objElement.tagname = "input") Then
                    If (objElement.Attribute("role") = "textbox") Or _
                       (objElement.Attribute("type") = "text") Then
                        If Found_TVL = "Y" Then
                            state = 6 'need to handle TVL field next
                        Else
                            state = 7 'no TVL field - skip it
                        End If
                        'Status Bar
                        sBar = StatusBar_Draw2("Time Entry", Round(((MajorElements + ((i * 12) + 5)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (i + 1) & " of " & TotalRows, Round((5 / 12) * 100, 0))
                        objElement.SendKeys ("") 'set focus to this element
                        If Not (objElement.Value = UCase(theShift)) Then
                            If Not (objElement.Value = "") Then
                                objElement.Clear
                            End If
                            Call SB_Wait(driver, SingleDelay)
                            objElement.SendKeys (theShift)
                            Call SB_Wait(driver, SingleDelay)
                        End If
                    End If
                End If
            ElseIf state = 6 Then   'next input field is TVL
                If (objElement.tagname = "div") Then
                    If (InStr(objElement.Attribute("class"), "sapMCbMark") > 0) Then
                        state = 7
                        j = 0
                        If Not ((InStr(objElement.Attribute("class"), "sapMCbMarkChecked") > 0) = bTVL) Then
                            objElement.Click
                            Call SB_Wait(driver, SingleDelay)
                        End If
                    End If
                End If
            ElseIf state = 7 Then   'next input fields are Hours
                If (objElement.tagname = "input") Then
                    If (objElement.Attribute("role") = "textbox") Or _
                       (objElement.Attribute("type") = "text") Then
                        objElement.SendKeys ("") 'set focus to this element
                        'TEMPO allows tenths of hours - compare both values as numbers rounded to 1 decimal place
                        If Not (Round(Val(objElement.Value), 1) = Round(Val(theHours(j)), 1)) Then
                            If Not (objElement.Value = "") Then
                                objElement.Clear
                            End If
                            Call SB_Wait(driver, SingleDelay)
                            objElement.SendKeys (theHours(j))
                            Call SB_Wait(driver, SingleDelay)
                        End If
                        'Status Bar
                        sBar = StatusBar_Draw2("Time Entry", Round(((MajorElements + ((i * 12) + 6 + j)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (i + 1) & " of " & TotalRows, Round(((6 + j) / 12) * 100, 0))
                        j = j + 1
                        If j > theHoursUBound Then
                            state = 8
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    Loop Until (Not (StartOver))
    
    If state <> 8 Then
        Excel_Activate
        result = MsgBox("Unable to enter Charge Object for row " & rowIndex & " in TEMPO.", vbExclamation)
        SB_Finish
    End If
End Sub

Sub SB_VerifyChargeObj_TEMPO_Init(driver As WebDriver)
' Initializes private global variables to keep track of HTML elements

    SB_numElements = driver.FindElementsByXPath("//span[@title='Delete Line'] | //span[@title='Add Line'] | //input").Count
    SB_elementIndex = 1
    SB_rowCount = 0
    SB_mismatch = ""

End Sub

Function SB_VerifyChargeObj_TEMPO(driver As WebDriver, rowIndex As Integer, theChargeObj As String, theExt As String, theShift As String, theHours() As String, theHoursUBound As Integer, theWPM As String, MajorElements As Integer, loopCount As Integer, TotalElements As Integer, TotalRows As Integer, Optional ByVal sTVL As String = "") As Boolean
' Verifies the entries for the Charge Object theValue in row rowIndex
'
    Dim objElement As WebElement
    Dim evt As Object
    Dim j As Integer
    Dim state As Integer
    Dim result As Integer
    Dim matches As Boolean
    Dim bTVL As Boolean
    Dim CellNum As Integer
    Dim sBar As Boolean
    
    CellNum = -1

    Call SB_Check_WD_Job(driver)    'check for WD Job field (sets Found_WD_Job to "Y" or "N")
    Call SB_Check_TVL(driver)       'check for TVL field (sets Found_TVL to "Y" or "N")
    If theShift = "" Then
        theShift = "1"
    End If
    If sTVL = "" Then
        bTVL = False
    Else
        bTVL = True
    End If
    state = 0
    matches = True
    If Debug_Warn Then Debug.Print SB_numElements & " elements"
    For SB_elementIndex = SB_elementIndex To SB_numElements
        Set objElement = driver.FindElementsByXPath("//span[@title='Delete Line'] | //span[@title='Add Line'] | //input")(SB_elementIndex)
        On Error Resume Next
            CellNum = Right(objElement.Attribute("tabindex"), 2)
        On Error GoTo 0
        If Debug_Warn Then
            Debug.Print objElement.tagname, objElement.Attribute("title")
            Debug.Print "Max Length - ", objElement.Attribute("maxlength")
            Debug.Print "Tab Index - ", objElement.Attribute("tabindex")
            Debug.Print "Cell Number", CellNum
        End If
        If CellNum = 27 And state = 5 Then ' Qualify the WPM field for only the line being entered/checked
            state = 4
        End If
        If state = 0 Then   'look for span with title "Delete Line"
            If (objElement.tagname = "span") Then
                If (objElement.Attribute("title") = "Delete Line") Then
                    If SB_rowCount = rowIndex Then
                        state = 1
                        'Status Bar
                        sBar = StatusBar_Draw2("Time Entry Check", Round(((MajorElements + ((rowIndex * 12) + 1)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (rowIndex + 1) & " of " & TotalRows, Round((1 / 12) * 100, 0))
                   End If
                    SB_rowCount = SB_rowCount + 1
                End If
            End If
        ElseIf state = 1 Then   'look for input field for the charge object: tagName "INPUT" with role "textbox"
            If (objElement.tagname = "input") Then
                If (objElement.Attribute("role") = "textbox") Or _
                   (objElement.Attribute("type") = "text") Then
                    state = 2
                    'Status Bar
                    sBar = StatusBar_Draw2("Time Entry Check", Round(((MajorElements + ((rowIndex * 12) + 2)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (rowIndex + 1) & " of " & TotalRows, Round((2 / 12) * 100, 0))
                    objElement.SendKeys ("") 'set focus to this element (so user knows something is happening)
                    If Not (objElement.Value = UCase(theChargeObj)) Then
                        matches = False
                        SB_mismatch = "Mismatch: Charge Object " & objElement.Value & " <> " & UCase(theChargeObj)
                        If Debug_Warn Then Debug.Print SB_mismatch
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 2 Then   'next input field is Ext
            If (objElement.tagname = "input") Then
                If (objElement.Attribute("role") = "textbox") Or _
                   (objElement.Attribute("type") = "text") Then
                    If Found_WD_Job = "Y" Then
                        state = 3 'need to handle WD Job field next
                    Else
                        'no WD Job field - skip it
                        If Found_WPM = "Y" Then
                            state = 4 'need to handle WPM field next
                        Else
                            state = 5 'no WPM field - skip it
                        End If
                    End If
                    'Status Bar
                    sBar = StatusBar_Draw2("Time Entry Check", Round(((MajorElements + ((rowIndex * 12) + 3)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (rowIndex + 1) & " of " & TotalRows, Round((3 / 12) * 100, 0))
                    If Not (objElement.Value = UCase(theExt)) Then
                        matches = False
                        SB_mismatch = "Mismatch: Ext " & objElement.Value & " <> " & UCase(theExt)
                        If Debug_Warn Then Debug.Print SB_mismatch
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 3 Then   'next input field is WD Job
            If (objElement.tagname = "input") Then
                If (objElement.Attribute("role") = "textbox") Or _
                   (objElement.Attribute("type") = "text") Then
                    If Found_WPM = "Y" Then
                        state = 4 'need to handle WPM field next
                    Else
                        state = 5 'no WPM field - skip it
                    End If
                    ' We don't have a field for WD Job
                    '  set WD Job to blank, let user enter the correct code later if needed
                    If Not (objElement.Value = "") Then
                        matches = False
                        SB_mismatch = "Mismatch: WD Job " & objElement.Value & " not empty"
                        If Debug_Warn Then Debug.Print SB_mismatch
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 4 Then   'next input field is WPM
            If (objElement.tagname = "input") Then
                If (objElement.Attribute("role") = "textbox") Or _
                   (objElement.Attribute("type") = "text") Then
                    state = 5
                    'Status Bar
                     sBar = StatusBar_Draw2("Time Entry Check", Round(((MajorElements + ((rowIndex * 12) + 4)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (rowIndex + 1) & " of " & TotalRows, Round((4 / 12) * 100, 0))
                    ' UpTEMPO doesn't have a field for WPM
                    '  set WPM to blank, let user enter the correct code later if needed
                    If Not (objElement.Value = UCase(theWPM)) Then
                        matches = False
                        SB_mismatch = "Mismatch: WPM " & objElement.Value & " not empty"
                        If Debug_Warn Then Debug.Print SB_mismatch
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 5 Then   'next input field is Shift
            If (objElement.tagname = "input") Then
                If (objElement.Attribute("role") = "textbox") Or _
                   (objElement.Attribute("type") = "text") Then
                    If Found_TVL = "Y" Then
                        state = 6 'need to handle TVL field next
                    Else
                        state = 7 'no TVL field - skip it
                    End If
                    'Status Bar
                    sBar = StatusBar_Draw2("Time Entry Check", Round(((MajorElements + ((rowIndex * 12) + 5)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (rowIndex + 1) & " of " & TotalRows, Round((5 / 12) * 100, 0))
                    If Not (objElement.Value = theShift) Then
                        matches = False
                        SB_mismatch = "Mismatch: Shift " & objElement.Value & " <> " & theShift
                        If Debug_Warn Then Debug.Print SB_mismatch
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 6 Then   'next input field is TVL
            If (objElement.tagname = "input") Then
                If (objElement.Attribute("type") = "checkbox") Then
                    state = 7
                    j = 0
                    If Not (objElement.IsSelected = bTVL) Then
                        matches = False
                        SB_mismatch = "Mismatch: TVL " & objElement.IsSelected & " <> " & bTVL
                        If Debug_Warn Then Debug.Print SB_mismatch
                        Exit For
                    End If
                End If
            End If
        ElseIf state = 7 Then   'next input fields are Hours
            If (objElement.tagname = "input") Then
                If (objElement.Attribute("role") = "textbox") Or _
                   (objElement.Attribute("type") = "text") Then
                    'TEMPO allows tenths of hours - compare both values as numbers rounded to 1 decimal place
                    If Not (Round(Val(objElement.Value), 1) = Round(Val(theHours(j)), 1)) Then
                        matches = False
                        SB_mismatch = "Mismatch: Hours " & objElement.Value & " <> " & theHours(j)
                        If Debug_Warn Then Debug.Print SB_mismatch
                    End If
                    'Status Bar
                    sBar = StatusBar_Draw2("Time Entry Check", Round(((MajorElements + ((rowIndex * 12) + 6 + j)) / TotalElements) * 100, 0), "Loop #" & loopCount + 1 & ", Row #" & (rowIndex + 1) & " of " & TotalRows, Round(((6 + j) / 12) * 100, 0))
                    j = j + 1
                    If j > theHoursUBound Then
                        state = 8
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    If Debug_Warn Then Debug.Print "matches", rowIndex, matches
    
    SB_VerifyChargeObj_TEMPO = matches
End Function

Sub SB_Check_WD_Job(driver As WebDriver)
' Checks for the "WD Job" field in TEMPO labor entry screen
'
    Dim objElement As WebElement
    Dim Found As Boolean
    
    If (Found_WD_Job = "Y") Or (Found_WD_Job = "N") Then
        'already checked and set variable - nothing more to do
    Else
        Found = False
        For Each objElement In driver.FindElementsByXPath("//label")
            If Debug_Warn Then Debug.Print objElement.tagname, objElement.Text
            If (UCase(Trim(objElement.Text)) = "WD JOB") Then
                Found = True
                Exit For
            End If
        Next
        If Found Then
            Found_WD_Job = "Y"
        Else
            Found_WD_Job = "N"
        End If
    End If
End Sub

Sub SB_Check_WPM(driver As WebDriver)
' Checks for the "WPM" field in TEMPO labor entry screen
'
    Dim objElement As WebElement
    Dim Found As Boolean
    
    If (Found_WPM = "Y") Or (Found_WPM = "N") Then
        'already checked and set variable - nothing more to do
    Else
        Found = False
        For Each objElement In driver.FindElementsByXPath("//bdi")
            If Debug_Warn Then Debug.Print objElement.tagname, objElement.Text
            If (UCase(Trim(objElement.Text)) = "WPM") Then
                Found = True
                Exit For
            End If
        Next
        If Found Then
            Found_WPM = "Y"
        Else
            Found_WPM = "N"
        End If
    End If
End Sub

Sub SB_Check_TVL(driver As WebDriver)
' Checks for the "TVL" field in TEMPO labor entry screen
'
    Dim objElement As WebElement
    Dim Found As Boolean
    
    If (Found_TVL = "Y") Or (Found_TVL = "N") Then
        'already checked and set variable - nothing more to do
    Else
        Found = False
        For Each objElement In driver.FindElementsByXPath("//label")
            If Debug_Warn Then Debug.Print objElement.tagname, objElement.Text
            If (UCase(Trim(objElement.Text)) = "TVL") Then
                Found = True
                Exit For
            End If
        Next
        If Found Then
            Found_TVL = "Y"
        Else
            Found_TVL = "N"
        End If
    End If
End Sub

Sub SB_EnterDaysOff_TEMPO(driver As WebDriver, dayName As String, theDate As String, dayOffValue As String)
' Enters the day off at position dayIndex (0 to 7)
'
    Dim objElement As WebElement
    Dim dayNumStr As String
    Dim state As Integer
    Dim result As Integer

    dayNumStr = Format(theDate, "d")
    state = 0
    For Each objElement In driver.FindElementsByXPath("//label | //bdi | //button")
        If Debug_Warn Then Debug.Print objElement.tagname, objElement.Text
        If state = 0 Then   'look for label with the day name: tagName "LABEL" with innerText "Fri", for example
            If (objElement.tagname = "label") Or _
               (objElement.tagname = "bdi") Then
                If (UCase(Trim(objElement.Text)) = UCase(dayName)) Then
                    state = 1
                End If
            End If
        ElseIf state = 1 Then   'next label element must have correct day number
            If (objElement.tagname = "label") Or _
               (objElement.tagname = "bdi") Then
                If (Trim(objElement.Text) = dayNumStr) Then
                    state = 2
                Else
                    state = 0
                End If
            End If
        ElseIf state = 2 Then   'find next button: tagName "BUTTON"
            If (objElement.tagname = "button") Then
                state = 3
                If (objElement.Text = "") Or _
                   (UCase(Trim(objElement.Text)) = UCase(dayName)) Then
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
        SB_Finish
    End If
End Sub

Function SB_VerifyDaysOff_TEMPO(driver As WebDriver, dayName As String, theDate As String, dayOffValue As String) As Boolean
' Verfies the entered day off at position dayIndex (0 to 7)
'
    Dim objElement As WebElement
    Dim dayNumStr As String
    Dim state As Integer
    Dim result As Integer
    Dim matches As Boolean

    dayNumStr = Format(theDate, "d")
    state = 0
    matches = False
    For Each objElement In driver.FindElementsByXPath("//label | //bdi | //button")
        If Debug_Warn Then Debug.Print objElement.tagname, objElement.Text
        If state = 0 Then   'look for label with the day name: tagName "LABEL" with innerText "Fri", for example
            If (objElement.tagname = "label") Or _
               (objElement.tagname = "bdi") Then
                If (UCase(Trim(objElement.Text)) = UCase(dayName)) Then
                    state = 1
                End If
            End If
        ElseIf state = 1 Then   'next label element must have correct day number
            If (objElement.tagname = "label") Or _
               (objElement.tagname = "bdi") Then
                If (Trim(objElement.Text) = dayNumStr) Then
                    state = 2
                Else
                    state = 0
                End If
            End If
        ElseIf state = 2 Then   'find next button: tagName "BUTTON"
            If (objElement.tagname = "button") Then
                state = 3
                If (objElement.Text = "") Or _
                   (UCase(Trim(objElement.Text)) = UCase(dayName)) Then
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
        SB_Finish
    End If
    
    SB_VerifyDaysOff_TEMPO = matches
End Function

Function SB_GetWEDate_TEMPO(driver As WebDriver) As Date
' Returns the current Week Ending date from TEMPO
' Returns date of 0 if the current Week Ending date is not found
'
    Dim objElement As WebElement
    Dim theDate As Date
    Dim state As Integer
    Dim result As Integer
    
    theDate = 0
    state = 0
    'If there is an error (because the page is not fully loaded), exit this function
    On Error GoTo Exit_WEDate
    'Look through all elements
    For Each objElement In driver.FindElementsByXPath("//label | //bdi | //input")
        If Debug_Warn Then Debug.Print objElement.tagname, objElement.Text
        If state = 0 Then   'look for Payroll W/E label: tagName "LABEL" with innerText "Payroll W/E:"
            If (objElement.tagname = "label") Or _
               (objElement.tagname = "bdi") Then
                If (Trim(objElement.Text) = "Payroll W/E:") Then
                    state = 1
                End If
            End If
        ElseIf state = 1 Then   'grab next input field to get date: tagName "INPUT"
            If (objElement.tagname = "input") Then
                state = 2
                theDate = CDate(Trim(objElement.Value))
                Exit For
            End If
        End If
    Next

Exit_WEDate:
    SB_GetWEDate_TEMPO = theDate
End Function

Sub SB_TimeEntry_TEMPO(driver As WebDriver)
' Go to the Time Entry screen in TEMPO
'
    Dim waitTime As Integer
    Dim objElement As WebElement
    
    'Navigate to Time Entry
    
    'check for "Time Entry" link
    Set objElement = driver.FindElementByXPath("//*[contains(text(),'Time Entry')]", timeout:=15000)
    If driver.FindElementsByXPath("//*[contains(text(),'Time Entry')]", timeout:=15000).Count = 1 Then
        'found "Time Entry" link, click it to go to time entry page
        driver.FindElementByXPath("//*[contains(text(),'Time Entry')]").Click
    Else
        'did not find "Time Entry" link, fall back on URL & suffix
        driver.Get URL_TEMPO & Suffix_Time_Entry
    End If
    
    'Check whether page is loaded
    waitTime = 0
    Do While (SB_GetWEDate_TEMPO(driver) = 0) And (waitTime < timeout)
        'Wait a little longer
        Call SB_Wait(driver, DoubleDelay)
        waitTime = waitTime + DoubleDelay
    Loop
End Sub

Sub SB_Wait(driver As WebDriver, theDelaySeconds As Integer)
    DoEvents
    If theDelaySeconds >= 0 Then
        driver.Wait (theDelaySeconds * 250)
    End If
End Sub

Function SB_WhatsRunning(driver As WebDriver) As String
' Checks what TEMPO screen is showing in the browser, returns a string identifying the name
    Dim SB_URL As String
    Dim SB_Title As String
    
    SB_URL = driver.Url
    SB_Title = driver.Title
    If (SB_URL = URL_TEMPO) Then
        'check for login screen
        If SB_Title = "Logon" Then
            SB_WhatsRunning = "TEMPO Login Page"
        Else
            SB_WhatsRunning = "TEMPO Other Page"
        End If
        Exit Function
    End If
    If (SB_URL = URL_TEMPO & Suffix_Shell_Home) Then
        SB_WhatsRunning = "TEMPO Welcome Page"
        Exit Function
    End If
    If (SB_URL = URL_TEMPO & Suffix_Time_Entry) Then
        SB_WhatsRunning = "TEMPO Time Entry Page"
        Exit Function
    End If
    If (SB_URL = URL_LoggedOff) Then
        SB_WhatsRunning = "TEMPO Logged Off Page"
        Exit Function
    End If
    If Len(SB_URL) > Len(URL_TEMPO) Then
        If Left(SB_URL, Len(URL_TEMPO)) = URL_TEMPO Then
            SB_WhatsRunning = "TEMPO Other Page"
            Exit Function
        End If
    End If
    If (SB_URL = URL_Authentication) Then
        SB_WhatsRunning = "TEMPO Authentication Page"
        Exit Function
    End If
    SB_WhatsRunning = "NONE"
End Function

Sub SB_Finish()
'
' Quits
'
    End
End Sub

'Code Module SHA-512
'''8a9a642037096d5f0c53b12a6d267bb42979564789062e80f539088e345ea8eec6f9d1ec4be144ed05985d7b65167b07b12d6d03b9fa732e45352d6f5648d958