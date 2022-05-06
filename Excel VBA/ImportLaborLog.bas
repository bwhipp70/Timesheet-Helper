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
        thePos1 = InStr(UCase(accstring), "</TD>")
        If thePos1 > 0 Then 'found end of HTML cell
            endFound = True
            LineRemnant = Mid(accstring, thePos1 + 5)
            accstring = Left(accstring, thePos1 + 5)
        Else
            thePos2 = InStr(UCase(accstring), "<TD")
            If thePos2 > 0 Then 'found beginning of HTML cell (in case of sloppy HTML)
                endFound = True
                LineRemnant = Mid(accstring, thePos2 + 3)
                accstring = Left(accstring, thePos2 + 3)
            Else 'no end or beginning found - read another line
                Line Input #1, theLine
                accstring = accstring & " " & theLine
            End If
        End If
    Loop
    InputOneCell = Replace(accstring, """", "")
End Function
Function FindValueFromName(searchName, getLength)
    nameFound = False
    Do While (Not EOF(1)) And (Not nameFound)
        theLine = InputOneCell
        thePos = InStr(theLine, "name=" & searchName)
        If thePos > 0 Then
            nameFound = True
            thePos = InStr(theLine, "value=")
            If thePos > 0 Then
                FindValueFromName = Mid(theLine, thePos + 6, getLength)
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
        thePos = InStr(theLine, "name=" & searchName)
        If thePos > 0 Then
            nameFound = True
            thePos = InStr(thePos, theLine, "selected>")
            If thePos > 0 Then
                FindSelectedFromName = Mid(theLine, thePos + 9, getLength)
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
        If UCase(Right(theLine, 5)) = "</TD>" Then
            cellFound = True
        End If
    Loop
    thePos = InStr(theLine, "name=hrs")
    If cellFound And (thePos > 0) Then
        thePos = InStr(theLine, "value=")
        If thePos > 0 Then
            GetHours = Mid(theLine, thePos + 6, getLength)
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
        thePos = InStr(theLine, "name=perComments")
        If thePos > 0 Then
            cmtFound = True
            thePos = InStr(thePos, theLine, ">")
            tempStr = Mid(theLine, thePos + 1)
            endPos = InStr(tempStr, "<")
            If endPos > 0 Then
                tempStr = Left(tempStr, endPos - 1)
            Else
                tempStr = ""
            End If
            GetComment = tempStr
        End If
    Loop
    If Not cmtFound Then GetComment = ""
End Function

'Code Module SHA-512
'''d616b56b268ea28440dc16af0964028301f0f2692fb7a9c103828c25d98c90842f0e8acf3697731017b4412e0a02213c364a8f17918da40c0bd048c9d861d0f7