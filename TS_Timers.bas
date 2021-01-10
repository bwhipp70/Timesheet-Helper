Option Explicit

' Reference http://www.nullskull.com/a/1479/identifying-which-formulas-in-excel-are-slowing-down-workbook-recalaculation.aspx

' 0. Enable the Developer Tab -> File -> Options -> Customize Ribbon -> Main Tabs -> Check Developer
' 1.  Create a worksheet called "ExecutionTimes"
' 2.  Paste this entire file into a VBA module
' 3.  Recommend creating a button to call timeallsheets and timeonesheet subroutines



'32 bit Declarations Below
'********************************************************************
' Private Declare Function getFrequency Lib "kernel32" _
 Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
' Private Declare Function getTickCount Lib "kernel32" _
 Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
'********************************************************************

'64 bit Declaractions Below
'********************************************************************
' Taken from http://www.jkp-ads.com/articles/apideclarations.asp

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
 Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long

Private Declare PtrSafe Function getTickCount Lib "kernel32" _
 Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
'********************************************************************

Function timeSheet(ws As Worksheet, routput As Range) As Range
 Dim ro As Range
 Dim c As Range, ct As Range, rt As Range, u As Range

 ws.Activate
Set u = ws.UsedRange
Set ct = u.Resize(1)
Set ro = routput

For Each c In ct.Columns
 Set ro = ro.Offset(1)
 Set rt = c.Resize(u.Rows.Count)
 rt.Select
 ro.Cells(1, 1).Value = rt.Worksheet.Name & "!" & rt.Address
 ro.Cells(1, 2) = shortCalcTimer(rt, False)
 Next c
 Set timeSheet = ro

End Function

Sub timeallsheets()
 Call timeloopSheets
End Sub

Sub timeloopSheets(Optional wsingle As Worksheet)

 Dim ws As Worksheet, ro As Range, rAll As Range
 Dim rKey As Range, r As Range, rSum As Range
 Const where = "ExecutionTimes!a1"

 Set ro = Range(where)
 ro.Worksheet.Cells.ClearContents
 Set rAll = ro
 'headers
 rAll.Cells(1, 1).Value = "address"
 rAll.Cells(1, 2).Value = "time"

If wsingle Is Nothing Then
' all sheets
For Each ws In Worksheets
Set ro = timeSheet(ws, ro)
Next ws
Else
' or just a single one
 Set ro = timeSheet(wsingle, ro)
End If

'now sort results, if there are any

If ro.Row > rAll.Row Then
Set rAll = rAll.Resize(ro.Row - rAll.Row + 1, 2)
Set rKey = rAll.Offset(1, 1).Resize(rAll.Rows.Count - 1, 1)
' sort highest to lowest execution time
With rAll.Worksheet.Sort
 .SortFields.Clear

 .SortFields.Add Key:=rKey, _
 SortOn:=xlSortOnValues, Order:=xlDescending, _
 DataOption:=xlSortNormal

 .SetRange rAll
 .Header = xlYes
 .MatchCase = False
 .Orientation = xlTopToBottom
 .SortMethod = xlPinYin
 .Apply
End With
'  sum times
Set rSum = rAll.Cells(1, 3)
 rSum.Formula = "=sum(" & rKey.Address & ")"
' %ages formulas
For Each r In rKey.Cells
 r.Offset(, 1).Formula = "=" & r.Address & "/" & rSum.Address
 r.Offset(, 1).NumberFormat = "0.00%"
 Next r

 End If
 rAll.Worksheet.Activate

End Sub

Function shortCalcTimer(rt As Range, Optional bReport As Boolean = True) As Double
 Dim dTime As Double
 Dim sCalcType As String
 Dim lCalcSave As Long
 Dim bIterSave As Boolean
'
On Error GoTo Errhandl


' Save calculation settings.
 lCalcSave = Application.Calculation
 bIterSave = Application.Iteration
If Application.Calculation <> xlCalculationManual Then
 Application.Calculation = xlCalculationManual
End If

' Switch off iteration.
If Application.Iteration <> False Then
 Application.Iteration = False
End If

' Get start time.
 dTime = MicroTimer
If Val(Application.Version) >= 12 Then
 rt.CalculateRowMajorOrder
Else
 rt.Calculate
End If


' Calc duration.
 sCalcType = "Calculate " & CStr(rt.Count) & _
 " Cell(s) in Selected Range: " & rt.Address
 dTime = MicroTimer - dTime
On Error GoTo 0

 dTime = Round(dTime, 5)
 If bReport Then
 MsgBox sCalcType & " " & CStr(dTime) & " Seconds"
End If

 shortCalcTimer = dTime

Finish:

' Restore calculation settings.
 If Application.Calculation <> lCalcSave Then
 Application.Calculation = lCalcSave
End If
If Application.Iteration <> bIterSave Then
 Application.Calculation = bIterSave
End If
 Exit Function
Errhandl:
On Error GoTo 0
 MsgBox "Unable to Calculate " & sCalcType, _
 vbOKOnly + vbCritical, "CalcTimer"
 GoTo Finish

End Function
'
Function MicroTimer() As Double
'

' Returns seconds.
'
 Dim cyTicks1 As Currency
 Static cyFrequency As Currency
'
 MicroTimer = 0

 ' Get frequency.
 If cyFrequency = 0 Then getFrequency cyFrequency

 ' Get ticks.
 getTickCount cyTicks1

' Seconds
If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function


Sub timeonesheet()
' Create a Button for this on ExecutionTimes worksheet
 
' Insert the name of the Worksheet that you want to time.
 
 Dim SheetTimer As String
 
 Sheets("ExecutionTimes").Select

 SheetTimer = Range("K8").Value
 
 Call timeloopSheets(Worksheets(SheetTimer))

End Sub

' *************************************************
' * Timer code for performance measuring
' * Taken from:  https://msdn.microsoft.com/en-us/library/office/ff700515(v=office.14).aspx
' *************************************************

'Private Declare Function getFrequency Lib "kernel32" _
'Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
'Private Declare Function getTickCount Lib "kernel32" _
'Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long


'Function MicroTimer() As Double
'

' Returns seconds.
'    Dim cyTicks1 As Currency
'    Static cyFrequency As Currency
    '
'    MicroTimer = 0

' Get frequency.
'    If cyFrequency = 0 Then getFrequency cyFrequency

' Get ticks.
'    getTickCount cyTicks1

' Seconds
'    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
'End Function


Sub RangeTimer()
    DoCalcTimer 1
End Sub
Sub SheetTimer()
    DoCalcTimer 2
End Sub
Sub RecalcTimer()
    DoCalcTimer 3
End Sub
Sub FullcalcTimer()
    DoCalcTimer 4
End Sub

Sub DoCalcTimer(jMethod As Long)
    Dim dTime As Double
    Dim dOvhd As Double
    Dim oRng As Range
    Dim oCell As Range
    Dim oArrRange As Range
    Dim sCalcType As String
    Dim lCalcSave As Long
    Dim bIterSave As Boolean
    '
    On Error GoTo Errhandl

' Initialize
    dTime = MicroTimer

    ' Save calculation settings.
    lCalcSave = Application.Calculation
    bIterSave = Application.Iteration
    If Application.Calculation <> xlCalculationManual Then
        Application.Calculation = xlCalculationManual
    End If
    Select Case jMethod
    Case 1

        ' Switch off iteration.

        If Application.Iteration <> False Then
            Application.Iteration = False
        End If
        
        ' Max is used range.

        If Selection.Count > 1000 Then
            Set oRng = Intersect(Selection, Selection.Parent.UsedRange)
        Else
            Set oRng = Selection
        End If

        ' Include array cells outside selection.

        For Each oCell In oRng
            If oCell.HasArray Then
                If oArrRange Is Nothing Then
                    Set oArrRange = oCell.CurrentArray
                End If
                If Intersect(oCell, oArrRange) Is Nothing Then
                    Set oArrRange = oCell.CurrentArray
                    Set oRng = Union(oRng, oArrRange)
                End If
            End If
        Next oCell

        sCalcType = "Calculate " & CStr(oRng.Count) & _
            " Cell(s) in Selected Range: "
    Case 2
        sCalcType = "Recalculate Sheet " & ActiveSheet.Name & ": "
    Case 3
        sCalcType = "Recalculate open workbooks: "
    Case 4
        sCalcType = "Full Calculate open workbooks: "
    End Select

' Get start time.
    dTime = MicroTimer
    Select Case jMethod
    Case 1
        If Val(Application.Version) >= 12 Then
            oRng.CalculateRowMajorOrder
        Else
            oRng.Calculate
        End If
    Case 2
        ActiveSheet.Calculate
    Case 3
        Application.Calculate
    Case 4
        Application.CalculateFull
    End Select

' Calculate duration.
    dTime = MicroTimer - dTime
    On Error GoTo 0

    dTime = Round(dTime, 5)
    MsgBox sCalcType & " " & CStr(dTime) & " Seconds", _
        vbOKOnly + vbInformation, "CalcTimer"

Finish:

    ' Restore calculation settings.
    If Application.Calculation <> lCalcSave Then
         Application.Calculation = lCalcSave
    End If
    If Application.Iteration <> bIterSave Then
         Application.Calculation = bIterSave
    End If
    Exit Sub
Errhandl:
    On Error GoTo 0
    MsgBox "Unable to Calculate " & sCalcType, _
        vbOKOnly + vbCritical, "CalcTimer"
    GoTo Finish
End Sub

