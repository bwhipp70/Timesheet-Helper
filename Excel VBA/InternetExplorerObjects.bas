Option Explicit
'
' To get ShellWindows: go to the Tools menu at the top of the VBA editor,
'  click on References in the menu, and scroll down the long list to find
'  the “Microsoft Internet Controls” reference. Click the checkbox to the
'  left of it, and then click OK.
'
' URL for TEMPO
'
Public URL_TEMPO As String
Const Suffix_Time_Entry = "#ZTPOTIMESHEET2-record"
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

    Call IE_GetUserValues(CallingSheet)
    
    'Find or open TEMPO in Internet Explorer
    Set objIE = IE_Find_Or_Open_TEMPO()
    Call IE_Wait_Until_Done(objIE)
    
    'go to Attendance & Labor Input page
    If IE_WhatsRunning = "TEMPO Home Page" Then
        IE_TimeEntry_TEMPO
    ElseIf IE_WhatsRunning = "TEMPO Other Page" Then
        IE_TimeEntry_TEMPO
    End If
'    'check for Logon page, perform logon if necessary
'    If IE_WhatsRunning = "STARS Logon Page" Then
'        Call IE_Logon_STARS(CallingSheet, False)
'    End If
'    'check for "Signed off, sign on again" page
'    If IE_WhatsRunning = "STARS Sign Off Sign On Page" Then
'        Call IE_Logon_STARS(CallingSheet, True)
'    End If
        
    'check for Attendance & Labor Input page and input labor
    If IE_WhatsRunning = "TEMPO Time Entry Page" Then
        'Get the current web browser session
        Set objIE = IE_Find_TEMPO()
        'Bring it to the front
        Call IE_Activate(objIE)
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
        End If
        'check whether week ending date in TEMPO matches this week's date
        If (CallingSheet = Labor_Flex980_ShName) Or _
            (CallingSheet = Labor_Flex980_2weeks_ShName) Then
            MMDDYYYYstr = IE_GetWEDate_TEMPO(objIE)
            If (CallingSheet = Labor_Flex980_ShName) Then
                WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
            ElseIf (CallingSheet = Labor_Flex980_2weeks_ShName) Then
                WEdate = Sheets(CallingSheet).Range("BH10").Value + 2
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
                    Call IE_EnterChargeObj_TEMPO(objIE, iEntries, Sheets(CallingSheet).Cells(iRow, 3).Value, _
                        Sheets(CallingSheet).Cells(iRow, 5).Value, Sheets(CallingSheet).Cells(iRow, 6).Value, theHours)
                ElseIf CallingSheet = Labor_Flex980_2weeks_ShName Then
                    theHours(0) = Sheets(CallingSheet).Cells(iRow, 13).Value 'Fri
                    theHours(1) = Sheets(CallingSheet).Cells(iRow, 14).Value 'Sat
                    theHours(2) = Sheets(CallingSheet).Cells(iRow, 15).Value 'Sun
                    theHours(3) = Sheets(CallingSheet).Cells(iRow, 16).Value 'Mon
                    theHours(4) = Sheets(CallingSheet).Cells(iRow, 17).Value 'Tue
                    theHours(5) = Sheets(CallingSheet).Cells(iRow, 18).Value 'Wed
                    theHours(6) = Sheets(CallingSheet).Cells(iRow, 19).Value 'Thu
                    theHours(7) = Sheets(CallingSheet).Cells(iRow, 20).Value 'Fri
                    Call IE_EnterChargeObj_TEMPO(objIE, iEntries, Sheets(CallingSheet).Cells(iRow, 3).Value, _
                        Sheets(CallingSheet).Cells(iRow, 5).Value, Sheets(CallingSheet).Cells(iRow, 6).Value, theHours)
                End If
                iEntries = iEntries + 1
            End If
            iRow = iRow + 1
        Loop Until iRow > LastEntryRow
        'delete extra labor lines at end
        Call IE_DeleteRows_TEMPO(objIE, iEntries)
        'unable to click the save button at this time
        'Call IE_Save_TEMPO(objIE)
        'pause briefly before returning to Excel
        Call IE_Wait(1)
        Excel_Activate
        If (CompletedDialogX <> "") Then
            result = MsgBox("Labor entry completed!" & Chr(10) & Chr(10) & _
                "Remember to review the labor and click the Save button in TEMPO.", vbInformation)
        End If
    Else
        'not at Attendance & Labor Input page
        Excel_Activate
        result = MsgBox("Unable to get to TEMPO Time Entry page.", vbExclamation)
        IE_Finish
    End If
    
    Set objIE = Nothing

End Sub
Sub IE_GetUserValues(CallingSheet)
' Gets user values from Instructions page
'
    URL_TEMPO = Range("TEMPO_URL").Value
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
    
    'First, check each open window/tab for an active TEMPO session
    foundTEMPO = False
    For Each objIE In shellWins
    
        IE_TabURL = objIE.LocationURL
        
        If (IE_TabURL = URL_TEMPO) Then
            'Found a valid TEMPO URL
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
    
End Function
Sub IE_DeleteRows_TEMPO(objIE As Object, rowIndex As Integer)
' Deletes rows from rowIndex and beyond
'
    Dim objElement As Object
    Dim i As Integer
    Dim result As Integer
    Dim StartOver As Boolean

    Do
        i = 0
        StartOver = False
        'Debug.Print objIE.Document.Count
        For Each objElement In objIE.Document.all
            'Debug.Print objElement.tagName, objElement.ID
            If (objElement.tagName = "SPAN") Then
                If (objElement.Title = "Delete Line") Then
                    If i = rowIndex Then
                        objElement.Click
                        Call IE_Wait(DoubleDelay)
                        StartOver = True
                        Exit For
                    Else
                        i = i + 1
                    End If
                ElseIf (objElement.Title = "Add Line") Then
                    Exit For
                End If
            End If
        Next
    Loop Until (Not (StartOver))
    
End Sub
Sub IE_EnterChargeObj_TEMPO(objIE As Object, rowIndex As Integer, theChargeObj As String, theExt As String, theShift As String, theHours() As String)
' Enters Charge Object theValue in row rowIndex
'
    Dim objElement As Object
    Dim evt As Object
    Dim i, j As Integer
    Dim state As Integer
    Dim result As Integer
    Dim StartOver As Boolean

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
                    If (objElement.role = "textbox") Then
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
                    If (objElement.role = "textbox") Then
                        state = 3
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
            ElseIf state = 3 Then   'next input field is Shift
                If (objElement.tagName = "INPUT") Then
                    If (objElement.role = "textbox") Then
                        state = 4
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
            ElseIf state = 4 Then   'next input fields are Hours
                If (objElement.tagName = "INPUT") Then
                    If (objElement.role = "textbox") Then
                        objElement.Focus
                        If Not (objElement.Value = UCase(theHours(j))) Then
                            objElement.Value = theHours(j)
                            'objElement.Click
                            Set evt = objIE.Document.createEvent("HTMLEvents")
                            'Set evt = objIE.Document.createEvent("keyboardevent")
                            evt.initEvent "change", True, False
                            objElement.dispatchEvent evt
                            Call IE_SendKeys(objIE, "{TAB}")
                            'Call IE_Wait(SingleDelay)
                        End If
                        j = j + 1
                        If j > UBound(theHours) Then
                            state = 5
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    Loop Until (Not (StartOver))
    
    If state <> 5 Then
        Excel_Activate
        result = MsgBox("Unable to enter Charge Object for row " & rowIndex & " in TEMPO.", vbExclamation)
        IE_Finish
    End If
End Sub
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
    Debug.Print LBound(theHours), UBound(theHours)
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
            If (objElement.tagName = "LABEL") Then
                If (UCase(Trim(objElement.innerText)) = UCase(dayName)) Then
                    state = 1
                End If
            End If
        ElseIf state = 1 Then   'next label element must have correct day number
            If (objElement.tagName = "LABEL") Then
                If (Trim(objElement.innerText) = dayNumStr) Then
                    state = 2
                Else
                    state = 0
                End If
            End If
        ElseIf state = 2 Then   'find next button: tagName "BUTTON"
            If (objElement.tagName = "BUTTON") Then
                state = 3
                If objElement.textContent = "" Then 'button is On (textContent is empty)
                    If dayOffValue = "" Then 'desired value is On
                        'already On, nothing to do
                    Else 'desired value is Off
                        objElement.Click 'click button to change to Off
                    End If
                Else 'button is Off (textContent = "OFF")
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
            If (objElement.tagName = "LABEL") Then
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

    If state <> 2 Then
        Excel_Activate
        result = MsgBox("Unable to find Payroll W/E date in TEMPO.", vbExclamation)
        IE_Finish
    End If
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
        Debug.Print objElement.tagName, objElement.ID
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
    
    'Get the current web browser session
    Set objIE = IE_Find_TEMPO()
    'Bring it to the front
    Call IE_Activate(objIE)
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
    Application.Wait DateAdd("s", theDelaySeconds, Now)
End Sub
Sub IE_Wait_Until_Done(objIE As Object)
    Dim IEbusy
    Dim IEreadyState
    Dim theType As String
    Dim waitComplete As Boolean
    
    Debug.Print "IE_Wait_Until_Done"
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
            Debug.Print IEbusy, IEreadyState
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
Function IE_GetMessage_STARS(objIE As Object) As String
' Returns the error message on the entry screen in STARS
'
    Dim objElements As Object
    Dim theMessage As String

    theMessage = ""
    On Error Resume Next
    Set objElements = objIE.Document.getElementsByClassName("starsMessage")
    On Error GoTo 0
    If objElements Is Nothing Then
        'Unable to find class name "starsMessage", so use an alternate method to find the error message
        theMessage = IE_GetMessage_STARS_Alternate(objIE)
    Else
        If objElements.Length <> 0 Then
            theMessage = Trim(objElements.Item(0).innerText)
        End If
    End If
    IE_GetMessage_STARS = theMessage
End Function
Function IE_GetMessage_STARS_Alternate(objIE As Object) As String
' Returns the error message on the entry screen in STARS
'
    Dim objElement As Object
    Dim theStr As String
    Dim theMessage As String

    theMessage = ""
    'Look through all elements
    Debug.Print "IE_GetMessage_STARS_Alternate"
    For Each objElement In objIE.Document.getElementsByTagName("td")
        theStr = objElement.innerHTML
'        Debug.Print theStr
        'alternate method: error message is inside a class "starsBorder"
        If InStr(theStr, "class=starsBorder") > 0 Then
            theMessage = Trim(objElement.innerText)
            Debug.Print theMessage
            Exit For
        End If
    Next
    IE_GetMessage_STARS_Alternate = theMessage
End Function
Function IE_WhatsRunning()
' Checks what STARS screen is showing in the terminal session, returns a string:
    Dim objIE As Object
    Dim IE_TabURL As String
    Dim objElements As Object
    Dim theMessage As String
    
    Set objIE = IE_Find_TEMPO()
    Call IE_Wait_Until_Done(objIE)
    IE_TabURL = objIE.LocationURL
    If (IE_TabURL = URL_TEMPO) Then
        IE_WhatsRunning = "TEMPO Welcome Page"
        Exit Function
    End If
    If (IE_TabURL = URL_TEMPO & Suffix_Time_Entry) Then
        IE_WhatsRunning = "TEMPO Time Entry Page"
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
Sub IE_Activate(objIE As Object)
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
End Sub
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
    Application.SendKeys (theString)
End Sub
Sub IE_Finish()
'
' Quits
'
    End
End Sub

