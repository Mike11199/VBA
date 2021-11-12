Attribute VB_Name = "Windows_API_Functions"
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_SNAPSHOT = &H2C
Public Const VK_MENU = &H12


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems


Public Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Public Declare Function GetWindowPlacement Lib "user32" _
(ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function SetWindowPlacement Lib "user32" _
(ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function SetForegroundWindow Lib "user32" _
(ByVal hwnd As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function BringWindowToTop Lib "user32" _
(ByVal hwnd As Long) As Long

Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2

Public Type POINTAPI
X As Long
y As Long
End Type

Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Public Type WINDOWPLACEMENT
Length As Long
flags As Long
showCmd As Long
ptMinPosition As POINTAPI
ptMaxPosition As POINTAPI
rcNormalPosition As RECT
End Type

Public Function ActivateWindow(xhWnd&) As Boolean

Dim Result&, WndPlcmt As WINDOWPLACEMENT

With WndPlcmt

.Length = Len(WndPlcmt)
Result = GetWindowPlacement(xhWnd, WndPlcmt)

If Result Then

If .showCmd = SW_SHOWMINIMIZED Then
.flags = 0
.showCmd = SW_SHOWNORMAL
Result = SetWindowPlacement(xhWnd, WndPlcmt)
Else
Call SetForegroundWindow(xhWnd)
Result = BringWindowToTop(xhWnd)
End If

If Result Then ActivateWindow = True

End If

End With

End Function

Public Function DeActivateWindow(xhWnd&) As Boolean

Dim Result&, WndPlcmt As WINDOWPLACEMENT

With WndPlcmt

.Length = Len(WndPlcmt)
Result = GetWindowPlacement(xhWnd, WndPlcmt)

If Result Then
.flags = 0
.showCmd = SW_SHOWMINIMIZED
Result = SetWindowPlacement(xhWnd, WndPlcmt)
If Result Then DeActivateWindow = True
End If

End With

End Function


