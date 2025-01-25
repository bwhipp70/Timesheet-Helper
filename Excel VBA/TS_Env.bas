'*****************************
' Thanks to:
'   https://support.microsoft.com/en-us/office/-how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel-189a257e-ec1b-40f7-9195-56d82e673071
'   https://stackoverflow.com/questions/5300770/how-to-check-from-net-code-whether-trust-access-to-the-vba-project-object-mode
'   https://www.mrexcel.com/board/threads/checking-if-trust-access-to-visual-basic-project-is-ticked.659774/
'   https://www.toolbox.com/tech/programming/question/excel-delete-a-command-button-on-a-worksheet-using-vba-070510/
'   https://stackoverflow.com/questions/2784367/capture-output-value-from-a-shell-command-in-vba
'   https://stackoverflow.com/questions/61331254/finding-and-deleting-one-line-of-excel-vba-code
'   https://forum.ozgrid.com/forum/index.php?thread/53016-macro-remove-macros-from-buttons/
'   https://stackoverflow.com/questions/44349684/replace-text-in-code-module
'   https://stackoverflow.com/questions/19800184/vbcomponents-remove-doesnt-always-remove-module
'   https://stackoverflow.com/questions/57958163/how-to-remove-reference-programmatically
'
'*****************************

Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
'        If sLine <> "" Then s = s & sLine & vbCrLf
        If sLine <> "" Then s = s & sLine
    Wend

    ShellRun = s

End Function
Public Sub SetUpEnv()

    Dim Check1 As String
    Dim Part1 As String
    Dim Part2 As String
    Dim iMarker As Long
    Dim hTotal As String
    Dim hPart1 As String
    Dim OK As Boolean
    Dim CheckIt As String
        
    ' 4.04 - Added VBA in front of Chr$
    CheckIt = VBA.Chr$(119) & VBA.Chr$(104) & VBA.Chr$(111) & VBA.Chr$(97) & VBA.Chr$(109) & VBA.Chr$(105)
    Check1 = ShellRun(CheckIt)
    iMarker = VBA.InStr(Check1, "\")
    ' Debug.Print Check1 & " " & iMarker
    ' 4.04 - Added VBA in front of Left and LCase
    Part1 = VBA.Left(Check1, iMarker - 1)
    hTotal = SHA512(VBA.LCase(Check1), False)
    hPart1 = SHA512(VBA.LCase(Part1), False)

    OK = Env1("AA", hPart1)
    
    If Not OK Then
        If Debug_Warn Then Debug.Print "Not OK 1"
        NotOK
    Else
        If Debug_Warn Then Debug.Print "OK 1"
    End If
    
    OK = Env1("AB", hTotal)
    
    If Not OK Then
        If Debug_Warn Then Debug.Print "Not OK 2"
    Else
        If Debug_Warn Then Debug.Print "OK 2"
        'UnlockMe
        Test_UnlockProject
        ThisWorkbook.VBProject.VBE.MainWindow.Visible = False
    End If
    
    OK = Env1("AC", hTotal)
    
    If Not OK Then
        If Debug_Warn Then Debug.Print "OK 3"
    Else
        If Debug_Warn Then Debug.Print "Not OK 3"
        NotOK
    End If
    

End Sub
Private Function Env1(whereisit As String, hash As String) As Boolean

    Dim offset As Integer
    Dim Found As Boolean
        
    offset = 0
    Found = False
    
    Do While Worksheets("VBA Trust Warning").Range(whereisit & (50 + offset)).Value <> ""
        If Worksheets("VBA Trust Warning").Range(whereisit & (50 + offset)).Value = hash Then
            Found = True
            Exit Do
        Else
            offset = offset + 1
        End If
    Loop
    
    Env1 = Found

End Function
Private Sub UnlockMe()
    
    Dim VBProj As Object
    Set VBProj = ThisWorkbook.VBProject
    If VBProj.Protection <> 1 Then Exit Sub ' already unprotected
    'Debug.Print "Welcome Developer, Unlocking Macros."
    'MsgBox "Welcome Developer, Unlocking VBA Project."
    Set Application.VBE.ActiveVBProject = VBProj
    SendKeys "LM" & "~~"
    Application.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True).Execute

End Sub
Private Sub DeleteVBComponent(ByVal wb As Workbook, ByVal CompName As String)

'Disabling the alert message
Application.DisplayAlerts = False

'Ignore errors
On Error Resume Next

'Delete the component
wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(CompName)

On Error GoTo 0

'Enabling the alert message
Application.DisplayAlerts = True

End Sub
Private Sub DeleteLines()

' Be very careful of this routine!
' If you delete lines above a function that will run after this, you will get weird errors that don't make sense.
' You are better off modifying the lines to comments that you want to change to preserve the flow.

' Modified to not require the VBIDE Reference
    Dim FindWhat As String
    Dim SL As Long ' start line
    Dim EL As Long ' end line
    Dim SC As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean

    ' 4.04 - Added VBA in front of Chr$
    FindWhat = VBA.Chr$(39) & VBA.Chr$(33) & VBA.Chr$(64) & VBA.Chr$(35) & VBA.Chr$(36) & VBA.Chr$(37)

    With ActiveWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
        SL = 1
        EL = .CountOfLines
        SC = 1
        EC = 255
        Found = .Find(Target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
            EndLine:=EL, EndColumn:=EC, _
            wholeword:=True, MatchCase:=False, patternsearch:=False)
        If Found = True Then
            .DeleteLines StartLine:=SL, Count:=3
        End If
    End With

End Sub
Private Sub ModLines()

' Modified to not require the VBIDE Reference
    Dim FindWhat As String
    Dim SL As Long ' start line
    Dim EL As Long ' end line
    Dim SC As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean

    ' 4.04 - Added VBA in front of Chr$
    FindWhat = VBA.Chr$(39) & VBA.Chr$(33) & VBA.Chr$(64) & VBA.Chr$(35) & VBA.Chr$(36) & VBA.Chr$(37)

    With ActiveWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
        SL = 1
        EL = .CountOfLines
        SC = 1
        EC = 255
        Found = .Find(Target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
            EndLine:=EL, EndColumn:=EC, _
            wholeword:=True, MatchCase:=False, patternsearch:=False)
        If Found = True Then
            .ReplaceLine SL, "    '"
            .ReplaceLine SL + 1, "    '"
            .ReplaceLine SL + 2, "    '"
        End If
    End With

End Sub
Private Sub ModLines2()

' Modified to not require the VBIDE Reference
    Dim FindWhat As String
    Dim SL As Long ' start line
    Dim EL As Long ' end line
    Dim SC As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean

    FindWhat = "Sub LE_EnterLabor(CallingSheet)"

    With ActiveWorkbook.VBProject.VBComponents("LaborEntry").CodeModule
        SL = 1
        EL = .CountOfLines
        SC = 1
        EC = 255
        Found = .Find(Target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
            EndLine:=EL, EndColumn:=EC, _
            wholeword:=True, MatchCase:=False, patternsearch:=False)
        If Found = True Then
            .ReplaceLine SL + 2, "'"
            .ReplaceLine SL + 5, "'"
        End If
    End With

End Sub
Sub UnCheck_Selenium()
Dim Int1 As Integer
Dim ref As Reference

With ThisWorkbook.VBProject.References
    For Int1 = 1 To .Count
        If .Item(Int1).GUID = "{0277FC34-FD1B-4616-BB19-A9AABCAF2A70}" Then
            If Debug_Warn Then Debug.Print "Selenium Reference Found."
            Debug.Print .Item(Int1).Name
            Set ref = ThisWorkbook.VBProject.References("Selenium")
            ThisWorkbook.VBProject.References.Remove ref
            Exit Sub
        End If
    Next
End With

End Sub
Private Sub NotOK()

    Test_UnlockProject
    ThisWorkbook.VBProject.VBE.MainWindow.Visible = False
    
    Sheets("Configuration").Unprotect
    Sheets("VBA Trust Warning").Unprotect
    
    Sheets("Configuration").Shapes("Button 1").OnAction = vbNullString  'DevMode Toggle
    Sheets("Configuration").Shapes("Button 3").OnAction = vbNullString  'Clean for Distribution
    Sheets("Configuration").Shapes("Button 5").OnAction = vbNullString  'Educational Mode
    Sheets("VBA Trust Warning").Range("AA49:AC500").Value = ""
    
    Sheets("Configuration").Protect
    Sheets("VBA Trust Warning").Protect
    
    TS_DevMode_Off
    TS_EducationalMode_Off
    
    ' Modify/Delete Functions
    Call ModLines
    Call ModLines2
    Call DeleteVBComponent(ActiveWorkbook, "TS_LockUnlock")
    Call DeleteVBComponent(ActiveWorkbook, "TS_Env")
    Call UnCheck_Selenium
    
    ' Save the Workbook
    Application.OnTime Now(), "CloseAndSave"
    
End Sub

'Code Module SHA-512
'''7dcbc7977480e45b880cade684ae9c92704fe4853feb1f7a45c9de2555593418a77ab3ba6a39fa0d6e7f5958c868ea5c645e6ceff0cdaacc98bc7fe314d6e393