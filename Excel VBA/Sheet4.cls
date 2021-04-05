Public SheetTimer As String


Private Sub Worksheet_Activate()

Dim sht As Worksheet, txt As String

Sheets("ExecutionTimes").Select

Range("K8").Value = ""

cmbSheet.Clear

For Each sht In ActiveWorkbook.Worksheets
      cmbSheet.AddItem sht.Name
Next sht

cmbSheet.ListIndex = 0

End Sub

Public Sub cmbSheet_Change()

Sheets("ExecutionTimes").Select

Range("K8").Value = cmbSheet.Value
    
End Sub