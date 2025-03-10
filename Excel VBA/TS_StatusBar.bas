' https://stackoverflow.com/questions/5181164/how-can-i-create-a-progress-bar-in-excel-vba
' https://wellsr.com/vba/2017/excel/vba-application-statusbar-to-mark-progress/

Sub StatusBar_Updater()
Dim CurrentStatus As Integer
Dim NumberOfBars As Integer
Dim pctDone As Integer
Dim lastrow As Long, i As Long
lastrow = Range("a" & Rows.Count).End(xlUp).Row

'(Step 1) Display your Status Bar
NumberOfBars = 40
Application.StatusBar = "Name of Task: [" & VBA.Space(NumberOfBars) & "]"

For i = 1 To lastrow
'(Step 2) Periodically update your Status Bar
    CurrentStatus = Int((i / lastrow) * NumberOfBars)
    pctDone = Round(CurrentStatus / NumberOfBars * 100, 0)
    Application.StatusBar = "Name of Task: [" & VBA.String(CurrentStatus, "|") & _
                            VBA.Space(NumberOfBars - CurrentStatus) & "]" & _
                            " " & pctDone & "% Complete"
    DoEvents
    '--------------------------------------
    'the rest of your macro goes below here
    '
    '
    '--------------------------------------
Next i

'(Step 3) Clear the Status Bar when you're done
Application.StatusBar = ""

End Sub

Sub StatusBar_Clear()

' Clear the Status Bar
Application.StatusBar = ""

End Sub
Function StatusBar_Draw(TaskName As String, pctDone As Integer) As Boolean

' TaskName gets displayed prior to the StatusBar
' pctDone should be 0 to 100

Dim NumberOfBars As Integer
Dim CurrentStatus As Integer

NumberOfBars = 50

'Display your Status Bar
CurrentStatus = Round((pctDone / 100) * NumberOfBars, 0)

Application.StatusBar = TaskName & ": [" & VBA.String(CurrentStatus, "|") & _
                        VBA.Space(NumberOfBars - CurrentStatus) & "]" & _
                            " " & pctDone & "% Complete"
DoEvents

If Debug_Warn Then Debug.Print TaskName & " " & pctDone & " " & CurrentStatus

End Function
Function StatusBar_Draw2(PrimaryTaskName As String, PrimarypctDone As Integer, SecondaryTaskName As String, SecondarypctDone As Integer) As Boolean

' TaskName gets displayed prior to the StatusBar
' pctDone should be 0 to 100

Dim PrimaryNumberOfBars As Integer
Dim SecondaryNumberOfBars As Integer
Dim CurrentStatus As Integer

PrimaryNumberOfBars = 20
SecondaryNumberOfBars = 50

'Display your Status Bar
PrimaryCurrentStatus = Round((PrimarypctDone / 100) * PrimaryNumberOfBars, 0)
SecondaryCurrentStatus = Round((SecondarypctDone / 100) * SecondaryNumberOfBars, 0)

Application.StatusBar = PrimaryTaskName & ": [" & VBA.String(PrimaryCurrentStatus, "|") & _
                        VBA.Space(PrimaryNumberOfBars - PrimaryCurrentStatus) & "]" & _
                            " " & PrimarypctDone & "%; " & _
                        SecondaryTaskName & ": [" & VBA.String(SecondaryCurrentStatus, "|") & _
                        VBA.Space(SecondaryNumberOfBars - SecondaryCurrentStatus) & "]" & _
                            " " & SecondarypctDone & "%"
DoEvents

If Debug_Warn Then Debug.Print TaskName & " " & pctDone & " " & CurrentStatus

End Function
Sub StatusBarTest1()

End Sub

' https://stackoverflow.com/questions/18602979/how-to-give-a-time-delay-of-less-than-one-second-in-excel-vba

Dim i As Long
Dim t As Double
Dim result As Boolean

Const ms As Double = 0.000000011574

i = 0

For i = 1 To 75

    result = StatusBar_Draw("Test String", Round(CInt(i) * 100 / 75, 0))
'    Application.Wait Now + (ms * 1000)
    t = Timer
    Do Until Timer - t >= 0.1
        DoEvents
    Loop
    Debug.Print "Time Delay = " & Timer - t
    
Next i

StatusBar_Clear

End Sub
Sub StatusBarTest2()

' https://stackoverflow.com/questions/18602979/how-to-give-a-time-delay-of-less-than-one-second-in-excel-vba

Dim i, j As Long
Dim t As Double
Dim result As Boolean

Const ms As Double = 0.000000011574

i = 0
j = 0

For j = 1 To 15

    For i = 1 To 75
    
        result = StatusBar_Draw2("Overall % Complete", Round(CInt(j) * 100 / 15, 0), "SubTask Name", Round(CInt(i) * 100 / 75, 0))
    '    Application.Wait Now + (ms * 1000)
        t = Timer
        Do Until Timer - t >= 0.1
            DoEvents
        Loop
        Debug.Print "Time Delay = " & Timer - t
        
    Next i

Next j

StatusBar_Clear

End Sub

'Code Module SHA-512
'''ef2bfb1165459e3d934f1053f431a1824a7f0804486336728c3e8b6bf387295f3952a854c194db7571d678bf4424c2262579f945ad7e19de892d2be4d63ed74e