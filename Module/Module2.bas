Attribute VB_Name = "Module2"
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Option Explicit
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'declare for moving the form
Public Declare Function ReleaseCapture Lib "USER32" () As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

'Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
 Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
     
      
      Declare Function FindWindow _
       Lib "USER32" Alias "FindWindowA" _
       (ByVal lpClassName As String, _
       ByVal lpWindowName As String) _
       As Long

Global PicIndex As Long
Global Duration As Long
Global numFrames As Long
Global Flicker As Boolean
Global SeeFlag As String
Global XX1 As Single
Global YY1 As Single
Global XX2 As Single
Global YY2 As Single
Global strOut As String
Global Eye As Long

Public Sub Delay(HowLong As Date)
Dim TempTime As String
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
 If Topmost = True Then 'Make the window topmost
  SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 Else
  SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
  SetTopMostWindow = False
 End If
End Function





