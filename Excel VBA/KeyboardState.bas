'Macro Module: KeyboardState
' WRH 2016-11-04

'Workarounds for SendKeys bug adapted from:
'     https://support.microsoft.com/en-us/kb/179987
' and https://support.microsoft.com/en-us/kb/177674
' and the book "Visual Basic Programmer's Guide to the Win32 API" by Dan Appleman

Option Explicit

' Declare Type for API call:
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128   '  Maintenance string for PSS usage
End Type

' API declarations:

Private Declare PtrSafe Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
   (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare PtrSafe Sub keybd_event Lib "user32" _
   (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare PtrSafe Function GetKeyboardState Lib "user32" _
   (pbKeyState As Byte) As Long

Private Declare PtrSafe Function SetKeyboardState Lib "user32" _
   (lppbKeyState As Byte) As Long

' Constant declarations:
Const VK_NUMLOCK = &H90
Const VK_CAPITAL = &H14
Const VK_SCROLL = &H91
Const SCAN_CAPS = &H3A
Const SCAN_NUM = &H45
Const SCAN_SCROLL = &H46
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

' Current states - updated by Get_Keyboard_States()
Dim NumLockState As Byte
Dim CapsLockState As Byte
Dim ScrollLockState As Byte

' Get and store current keyboard states
Sub Get_Keyboard_States()
    Dim keys(0 To 255) As Byte
    
    GetKeyboardState keys(0)
    ' Num Lock
    NumLockState = keys(VK_NUMLOCK)
    ' Caps Lock
    CapsLockState = keys(VK_CAPITAL)
    ' Scroll Lock
    ScrollLockState = keys(VK_SCROLL)
End Sub

' Set keyboard states to stored values
Sub Set_Keyboard_States()
    Dim o As OSVERSIONINFO
    Dim keys(0 To 255) As Byte

    o.dwOSVersionInfoSize = Len(o)
    GetVersionEx o
    GetKeyboardState keys(0)
    If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then 'Win95/98
        keys(VK_NUMLOCK) = NumLockState
        keys(VK_CAPITAL) = CapsLockState
        keys(VK_SCROLL) = ScrollLockState
        SetKeyboardState keys(0)
    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then 'WinNT
        ' Num Lock
        If Not (keys(VK_NUMLOCK) = NumLockState) Then
            'Simulate key press
            keybd_event VK_NUMLOCK, SCAN_NUM, KEYEVENTF_EXTENDEDKEY Or 0, 0
            'Simulate key release
            keybd_event VK_NUMLOCK, SCAN_NUM, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
        ' Caps Lock
        If Not (keys(VK_CAPITAL) = CapsLockState) Then
            'Simulate key press
            keybd_event VK_CAPITAL, SCAN_CAPS, KEYEVENTF_EXTENDEDKEY Or 0, 0
            'Simulate key release
            keybd_event VK_CAPITAL, SCAN_CAPS, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
        ' Scroll Lock
        If Not (keys(VK_SCROLL) = ScrollLockState) Then
            'Simulate key press
            keybd_event VK_SCROLL, SCAN_SCROLL, KEYEVENTF_EXTENDEDKEY Or 0, 0
            'Simulate key release
            keybd_event VK_SCROLL, SCAN_SCROLL, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
    End If
End Sub



