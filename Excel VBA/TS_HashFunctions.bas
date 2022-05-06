Option Explicit

'**********************************
' Thanks to:
'   https://www.mrexcel.com/board/threads/close-the-visual-basic-editor-window.28529/
'   https://www.mrexcel.com/board/threads/lock-unlock-vbaprojects-programmatically-without-sendkeys.1136415/
'   https://stackoverflow.com/questions/31134582/protect-vba-project-using-vba
'   http://www.cpearson.com/Excel/VBE.aspx
'   https://stackoverflow.com/questions/5300770/how-to-check-from-net-code-whether-trust-access-to-the-vba-project-object-mode
'   https://www.mrexcel.com/board/threads/checking-if-trust-access-to-visual-basic-project-is-ticked.659774/
'   https://www.mrexcel.com/board/threads/protect-vba-macro-with-md5-check.510470/
'   https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA#:~:text=Hashes%20can%20be%20used%20as,password%20strings%20in%20their%20code.
'   https://www.excelforum.com/excel-programming-vba-macros/1233665-using-one-macro-to-change-lines-of-code-in-a-different-macro.html
'**********************************

Sub TestHash()
    'run this to test md5, sha1, sha2/256, sha384, sha2/512 with salt, or sha2/512
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim sIn As String, sOut As String, b64 As Boolean
    Dim sH As String, sSecret As String
    
    'insert the text to hash within the sIn quotes
    'and for selected procedures a string for the secret key
    sIn = ""
    sSecret = "" 'secret key for StrToSHA512Salt only
    
    'select as required
    'b64 = False   'output hex
    b64 = True   'output base-64
    
    'enable any one
    sH = MD5(sIn, b64)
    'sH = SHA1(sIn, b64)
    'sH = SHA256(sIn, b64)
    'sH = SHA384(sIn, b64)
    'sH = StrToSHA512Salt(sIn, sSecret, b64)
    'sH = SHA512(sIn, b64)
    
    'message box and immediate window outputs
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    MsgBox sH & vbNewLine & Len(sH) & " characters in length"
    
    'de-comment this block to place the hash in first cell of sheet1
'    With ThisWorkbook.Worksheets("Sheet1").Cells(1, 1)
'        .Font.Name = "Consolas"
'        .Select: Selection.NumberFormat = "@" 'make cell text
'        .Value = sH
'    End With

End Sub

Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
        
    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc
        
    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
        
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.Getbytes_4(sIn)
    bytes = oMD5.ComputeHash_2((TextToHash))
 
    If bB64 = True Then
       MD5 = ConvToBase64String(bytes)
    Else
       MD5 = ConvToHexString(bytes)
    End If
        
    Set oT = Nothing
    Set oMD5 = Nothing

End Function

Public Function SHA1(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
   
   'Test with empty string input:
    '40 Hex:   da39a3ee5e6...etc
    '28 Base-64:   2jmj7l5rSw0yVb...etc
    
    Dim oT As Object, oSHA1 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
            
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    bytes = oSHA1.ComputeHash_2((TextToHash))
        
    If bB64 = True Then
       SHA1 = ConvToBase64String(bytes)
    Else
       SHA1 = ConvToHexString(bytes)
    End If
            
    Set oT = Nothing
    Set oSHA1 = Nothing
    
End Function

Public Function SHA256(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    'Test with empty string input:
    '64 Hex:   e3b0c44298f...etc
    '44 Base-64:   47DEQpj8HBSa+/...etc
    
    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    bytes = oSHA256.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA256 = ConvToBase64String(bytes)
    Else
       SHA256 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA256 = Nothing
    
End Function

Public Function SHA384(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    'Test with empty string input:
    '96 Hex:   38b060a751ac...etc
    '64 Base-64:   OLBgp1GsljhM2T...etc
    
    Dim oT As Object, oSHA384 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    bytes = oSHA384.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA384 = ConvToBase64String(bytes)
    Else
       SHA384 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA384 = Nothing
    
End Function

Public Function SHA512(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    'Test with empty string input:
    '128 Hex:   cf83e1357eefb8bd...etc
    '88 Base-64:   z4PhNX7vuL3xVChQ...etc
    
    Dim oT As Object, oSHA512 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    bytes = oSHA512.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA512 = ConvToBase64String(bytes)
    Else
       SHA512 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA512 = Nothing
    
End Function

Function StrToSHA512Salt(ByVal sIn As String, ByVal sSecretKey As String, _
                           Optional ByVal b64 As Boolean = False) As String
    'Returns a sha512 STRING HASH in function name, modified by the parameter sSecretKey.
    'This hash differs from that of SHA512 using the SHA512Managed class.
    'HMAC class inputs are hashed twice;first input and key are mixed before hashing,
    'then the key is mixed with the result and hashed again.
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SecretKey() As Byte
    Dim bytes() As Byte
    
    'Test results with both strings empty:
    '128 Hex:    b936cee86c9f...etc
    '88 Base-64:   uTbO6Gyfh6pd...etc
    
    'create text and crypto objects
    Set asc = CreateObject("System.Text.UTF8Encoding")
    
    'Any of HMACSHAMD5,HMACSHA1,HMACSHA256,HMACSHA384,or HMACSHA512 can be used
    'for corresponding hashes, albeit not matching those of Managed classes.
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")

    'make a byte array of the text to hash
    bytes = asc.Getbytes_4(sIn)
    'make a byte array of the private key
    SecretKey = asc.Getbytes_4(sSecretKey)
    'add the private key property to the encryption object
    enc.Key = SecretKey

    'make a byte array of the hash
    bytes = enc.ComputeHash_2((bytes))
    
    'convert the byte array to string
    If b64 = True Then
       StrToSHA512Salt = ConvToBase64String(bytes)
    Else
       StrToSHA512Salt = ConvToHexString(bytes)
    End If
    
    'release object variables
    Set asc = Nothing
    Set enc = Nothing

End Function

Private Function ConvToBase64String(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
   
   Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Private Function ConvToHexString(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function
Function VBATrusted() As Boolean
On Error Resume Next
VBATrusted = (Application.VBE.VBProjects.Count) > 0
Exit Function
End Function
Private Sub ChangeTo_Workbook_Open()
If Not VBATrusted() Then
    MsgBox "No Access to VB Project" & vbLf & _
      "Please allow access in Trusted Sources" & vbLf & _
      "(File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust Access to the VBA Project object model)"
End If
End Sub
Private Sub UnlockMe()
    
    Dim VBProj As Object
    Set VBProj = ThisWorkbook.VBProject
    If VBProj.Protection <> 1 Then Exit Sub ' already unprotected
    Set Application.VBE.ActiveVBProject = VBProj
    SendKeys "LM" & "~~"
    Application.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True).Execute
    
    startTime = Timer
    Do
    Loop Until Timer - startTime >= 10

End Sub
Private Sub RegWrite_Workbook_Open()
    Dim aLibKey As String
    Dim WshShell
    
    Set WshShell = CreateObject("WScript.Shell")
    MsgBox WshShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM")
'    MsgBox WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM")

'    WshShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM", 1, "REG_DWORD"
'    WshShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM", 0, "REG_DWORD"

    MsgBox WshShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM")

End Sub
Public Sub CheckContents()
' This requires modifications to the Trust Settings
' File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust Access to the VBA Project object model

  Const MyName As String = "TS_HashFunctions"
 
  Dim obj As Object
  Dim cLines As Integer
  Dim sVBAcode As String
  Dim sMD5hash As String
  Dim iEndMarker As Long
 
  For Each obj In ActiveWorkbook.VBProject.VBComponents
    If obj.Type = 1 And obj.Name = MyName Then
      cLines = Application.VBE.ActiveVBProject.VBComponents(obj.Name).CodeModule.CountOfLines
      If cLines = 0 Then Stop
      sVBAcode = Application.VBE.ActiveVBProject.VBComponents(obj.Name).CodeModule.Lines(1, cLines)
      iEndMarker = InStr(sVBAcode, String(3, "'"))
      If iEndMarker = 0 Then Stop
      sMD5hash = Mid(sVBAcode, iEndMarker + 3)
      sVBAcode = Left(sVBAcode, iEndMarker)
      MsgBox "sMD5hash = " & sMD5hash & vbNewLine
      MsgBox "sVBACode = " & sVBAcode & vbNewLine
      ' sMD5hash contains the MD5 hash from the comment line at the end of this module[/COLOR]
      ' sVBAcode contains all of the VBA code before the MD5 comment line[/COLOR]
      ' now you just calculate the MD5 of the sVBAcode and compare it to sMD5hash[/COLOR]
    End If
  Next obj
 
End Sub
Public Function CheckAllHash() As Boolean

  Dim obj As Object
  Dim cLines As Integer
  Dim sVBAcode As String
  Dim sCodehash As String
  Dim iEndMarker As Long
  Dim sHash As String
  
  'UnlockMe
  Test_UnlockProject
  ThisWorkbook.VBProject.VBE.MainWindow.Visible = False
  
  For Each obj In ThisWorkbook.VBProject.VBComponents
    sVBAcode = ""
    sCodehash = "No Hash"
    If Debug_Warn Then Debug.Print obj.Name & " " & obj.Type
    If obj.Type = 1 Then
      cLines = Application.VBE.ActiveVBProject.VBComponents(obj.Name).CodeModule.CountOfLines
      If cLines = 0 Then Stop
      sVBAcode = Application.VBE.ActiveVBProject.VBComponents(obj.Name).CodeModule.Lines(1, cLines)
      iEndMarker = InStr(sVBAcode, String(3, "'"))
'      If iEndMarker = 0 Then Debug.Print "WARNING: No marker at the end of the VBA Code"
      If iEndMarker > 0 Then
        sCodehash = Mid(sVBAcode, iEndMarker + 3)
        sVBAcode = Left(sVBAcode, iEndMarker)
      End If
      sHash = SHA512(sVBAcode, False)
      If sHash = sCodehash Then
        If Debug_Warn Then Debug.Print "(MATCH) Hash for " & obj.Name & " = " & sHash
      Else
        If Debug_Warn Then Debug.Print "(FAIL) Hash for " & obj.Name & " = " & sHash; " vs. Code Hash of " & sCodehash
        CheckAllHash = False
        Exit Function
      End If
    End If
  Next obj

  CheckAllHash = True

End Function
Public Function FixHash(cm As String, sHash As String) As Boolean
' This requires modifications to the Trust Settings
' File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust Access to the VBA Project object model

' Modified to not require the VBIDE Reference
    Dim FindWhat As String
    Dim SL As Long ' start line
    Dim EL As Long ' end line
    Dim SC As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean

    FindWhat = Chr$(39) & Chr$(39) & Chr$(39)

    With ActiveWorkbook.VBProject.VBComponents(cm).CodeModule
        SL = 1
        EL = .CountOfLines
        SC = 1
        EC = 255
        Found = .Find(Target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
            EndLine:=EL, EndColumn:=EC, _
            wholeword:=False, MatchCase:=False, patternsearch:=False)
        If Found = True Then
            .DeleteLines StartLine:=SL, Count:=1
            .InsertLines SL, FindWhat & sHash
        Else
            Debug.Print "No Hash marker found it " & cm & ".  Adding one."
            Debug.Print "Re-run the Hash Routine again (added a comment)."
            .InsertLines .CountOfLines, "'Code Module SHA-512"
            .InsertLines .CountOfLines, FindWhat & sHash
            .DeleteLines .CountOfLines, Count:=1  'Inserting adds a CR/LF, need to delete it
        End If
    End With
 
End Function
Public Sub CheckIntegrity()
    Dim result As Boolean
    
    result = CheckAllHash
    
    If result Then
        MsgBox "Integrity check OK."
    Else
        MsgBox "Integrity check FAILED!"
    End If

End Sub
Public Sub ShowAndFixAllHash()
' This requires modifications to the Trust Settings
' File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust Access to the VBA Project object model

  Dim obj As Object
  Dim cLines As Integer
  Dim sVBAcode As String
  Dim sCodehash As String
  Dim iEndMarker As Long
  Dim sHash As String
  Dim test As Boolean
  
  For Each obj In ThisWorkbook.VBProject.VBComponents
    sVBAcode = ""
    sCodehash = "No Hash"
    If Debug_Warn Then Debug.Print obj.Name & " " & obj.Type
    If (obj.Type = 1 Or ((obj.Type = 100) And (obj.Name = "ThisWorkbook"))) Then
      cLines = Application.VBE.ActiveVBProject.VBComponents(obj.Name).CodeModule.CountOfLines
      If cLines = 0 Then Stop
      sVBAcode = Application.VBE.ActiveVBProject.VBComponents(obj.Name).CodeModule.Lines(1, cLines)
      iEndMarker = InStr(sVBAcode, String(3, "'"))
      If iEndMarker = 0 Then Debug.Print "WARNING: No marker at the end of the VBA Code"
      If iEndMarker > 0 Then
        sCodehash = Mid(sVBAcode, iEndMarker + 3)
        sVBAcode = Left(sVBAcode, iEndMarker)
      End If
      sHash = SHA512(sVBAcode, False)
      If sHash = sCodehash Then
        Debug.Print "(MATCH) Hash for " & obj.Name & " = " & sHash
      Else
        Debug.Print "(FAIL) Hash for " & obj.Name & " = " & sHash; " vs. Code Hash of " & sCodehash
        Debug.Print "Fixing the Hash Value..."
        test = FixHash(obj.Name, sHash)
      End If
    End If
  Next obj
 
End Sub
 
'Code Module SHA-512
'''981581a99812a8036cf7b84727ad3baa518aa962f138e075a1dfe21c42f7e545919fd0c7839c9882605279cab387c7f7309ccc7afdd59aa7344dd927d0803600