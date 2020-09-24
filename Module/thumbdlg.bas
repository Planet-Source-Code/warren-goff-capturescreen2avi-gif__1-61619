Attribute VB_Name = "thmbdialog"
Option Explicit

' Thumbdlg sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 2003 by Matt Hart
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code. Do not copy this sample to a collection, such as
' a CD-ROM archive. You may link directly to the original sample
' using "http://blackbeltvb.com/free/thumbdlg.htm"
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.

Public Const WH_CALLWNDPROC = 4
Public Const WM_CREATE = &H1

Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    Y As Long
    X As Long
    Style As Long
    lpszName As Long
    lpszClass As Long
    ExStyle As Long
End Type

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public hHook As Long
Public hToolbar As Long
Global EndFlag As Boolean, StepFlag As Boolean

Public Function AppHook(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim CWP As CWPSTRUCT
    CopyMemory CWP, ByVal lParam, Len(CWP)
    Select Case CWP.message
        Case WM_CREATE
            Dim a As String, l As Long
            a = Space$(128)
            l = GetClassName(CWP.hWnd, a, 128)
            Select Case Left$(a, l)
                Case "ToolbarWindow32"
                    hToolbar = CWP.hWnd
                Case "tooltips_class32"
                    If hToolbar Then
                        Panorama.Timer1.Enabled = True
                        AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
                        UnhookWindowsHookEx hHook
                        hHook = 0
                        Exit Function
                    End If
            End Select
    End Select
    AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
End Function
