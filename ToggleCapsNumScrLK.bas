Attribute VB_Name = "ToggleCapsNumScrLK"
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_CAPITAL = &H14
Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91

Public Function CapsLockOn() As Boolean
   CapsLockOn = GetKeyState(vbKeyCapital)
End Function

Public Function NumLockOn() As Boolean
    NumLockOn = GetKeyState(vbKeyNumlock)
End Function

Public Function ScrlLockOn() As Boolean
   ScrlLockOn = GetKeyState(vbKeyScrollLock)
End Function

Public Sub ToggleNumLock()
        keybd_event VK_NUMLOCK, 0, 1, 0
        keybd_event VK_NUMLOCK, 0, 2, 0
End Sub

Public Sub ToggleCapsLock()
        keybd_event VK_CAPITAL, 0, 1, 0
        keybd_event VK_CAPITAL, 0, 2, 0
End Sub


Public Sub ToggleScrollLock()
        keybd_event VK_SCROLL, 0, 1, 0
        keybd_event VK_SCROLL, 0, 2, 0
End Sub


