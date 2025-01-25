Option Explicit
'
Private LineRemnant As String
Dim accstring As String
'
Sub Import_Labor_Log_SAP_Flex980()
Dim result
    result = MsgBox("Sorry, Import SAP Labor Log is not yet available.", vbCritical)
End Sub
Sub Import_Labor_Log_SAP_Flex980_2weeks()
Dim result
    result = MsgBox("Sorry, Import SAP Labor Log is not yet available.", vbCritical)
End Sub
Function InputOneCell()
    endFound = False
    accstring = LineRemnant
    LineRemnant = ""
    Do While (Not EOF(1)) And (Not endFound)
        thePos1 = VBA.InStr(VBA.UCase(accstring), "</TD>")
        If thePos1 > 0 Then 'found end of HTML cell
            endFound = True
            LineRemnant = VBA.Mid(accstring, thePos1 + 5)
            accstring = VBA.Left(accstring, thePos1 + 5)
        Else
            thePos2 = VBA.InStr(VBA.UCase(accstring), "<TD")
            If thePos2 > 0 Then 'found beginning of HTML cell (in case of sloppy HTML)
                endFound = True
                LineRemnant = VBA.Mid(accstring, thePos2 + 3)
                accstring = VBA.Left(accstring, thePos2 + 3)
            Else 'no end or beginning found - read another line
                Line Input #1, theLine
                accstring = accstring & " " & theLine
            End If
        End If
    Loop
    InputOneCell = VBA.Replace(accstring, """", "")
End Function
Function FindValueFromName(searchName, getLength)
    nameFound = False
    Do While (Not EOF(1)) And (Not nameFound)
        theLine = InputOneCell
        thePos = VBA.InStr(theLine, "name=" & searchName)
        If thePos > 0 Then
            nameFound = True
            thePos = VBA.InStr(theLine, "value=")
            If thePos > 0 Then
                FindValueFromName = VBA.Mid(theLine, thePos + 6, getLength)
            Else
                FindValueFromName = ""
            End If
        End If
    Loop
    If Not nameFound Then FindValueFromName = ""
End Function
Function FindSelectedFromName(searchName, getLength)
    nameFound = False
    Do While (Not EOF(1)) And (Not nameFound)
        theLine = InputOneCell
        thePos = VBA.InStr(theLine, "name=" & searchName)
        If thePos > 0 Then
            nameFound = True
            thePos = VBA.InStr(thePos, theLine, "selected>")
            If thePos > 0 Then
                FindSelectedFromName = VBA.Mid(theLine, thePos + 9, getLength)
            Else
                FindSelectedFromName = ""
            End If
        End If
    Loop
    If Not nameFound Then FindSelectedFromName = ""
End Function
Function GetHours(getLength)
    cellFound = False
    'make sure we get a whole cell (to the ending "</TD>")
    Do While (Not EOF(1)) And (Not cellFound)
        theLine = InputOneCell
        If VBA.UCase(VBA.Right(theLine, 5)) = "</TD>" Then
            cellFound = True
        End If
    Loop
    thePos = VBA.InStr(theLine, "name=hrs")
    If cellFound And (thePos > 0) Then
        thePos = VBA.InStr(theLine, "value=")
        If thePos > 0 Then
            GetHours = VBA.Mid(theLine, thePos + 6, getLength)
        Else
            GetHours = ""
        End If
    Else
        GetHours = ""
    End If
End Function
Function GetComment()
    cmtFound = False
    Do While (Not EOF(1)) And (Not cmtFound)
        theLine = InputOneCell
        thePos = VBA.InStr(theLine, "name=perComments")
        If thePos > 0 Then
            cmtFound = True
            thePos = VBA.InStr(thePos, theLine, ">")
            tempStr = VBA.Mid(theLine, thePos + 1)
            endPos = VBA.InStr(tempStr, "<")
            If endPos > 0 Then
                tempStr = VBA.Left(tempStr, endPos - 1)
            Else
                tempStr = ""
            End If
            GetComment = tempStr
        End If
    Loop
    If Not cmtFound Then GetComment = ""
End Function

'Code Module SHA-512
'''5e38ec18ae74b075d85c035cde82f67368073b9332be840535430947f38392b973bdefc45d097a9c1dfb107ca6a2bb7a9533c147f76f93b038163a148f1eabe4