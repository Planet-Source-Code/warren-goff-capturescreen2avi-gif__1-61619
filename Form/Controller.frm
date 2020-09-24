VERSION 5.00
Begin VB.Form Controller 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Moose Controller"
   ClientHeight    =   3810
   ClientLeft      =   345
   ClientTop       =   765
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Controller.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Controller.frx":08CA
   ScaleHeight     =   3810
   ScaleWidth      =   5595
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4890
      Picture         =   "Controller.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Edit AVI Premiere-minus"
      Top             =   1050
      Width           =   705
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   390
      TabIndex        =   1
      Text            =   "Available Frames to edit"
      Top             =   3210
      Width           =   4860
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Ö"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Delete All Pictures"
      Top             =   3210
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H006C565C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Picture         =   "Controller.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "AVI to Animated GIF"
      Top             =   75
      Width           =   645
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   60
      ScaleHeight     =   2205
      ScaleWidth      =   3960
      TabIndex        =   27
      Top             =   930
      Width           =   3960
      Begin VB.Image imgMain 
         Height          =   2025
         Left            =   90
         Picture         =   "Controller.frx":2328
         Stretch         =   -1  'True
         Top             =   45
         Width           =   3855
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4110
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1230
      UseMaskColor    =   -1  'True
      Value           =   2  'Grayed
      Width           =   705
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00000000&
      Caption         =   "Append"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4110
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2640
      UseMaskColor    =   -1  'True
      Value           =   2  'Grayed
      Width           =   705
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Giffy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4875
      Picture         =   "Controller.frx":62AB
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Edit Animated GIF"
      Top             =   2040
      Width           =   705
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GIF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4890
      Picture         =   "Controller.frx":6B75
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "View Animated GIF"
      Top             =   75
      Width           =   705
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AVI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4890
      Picture         =   "Controller.frx":743F
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "View AVI"
      Top             =   570
      Width           =   705
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000006&
      Caption         =   "Pan"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   1565
      Picture         =   "Controller.frx":7D09
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   75
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5055
      Top             =   2745
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1575
      Picture         =   "Controller.frx":85D3
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Select Area and Start Capture"
      Top             =   75
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000006&
      Caption         =   "Step"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   825
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "You may Step Capture a frame at a time ("
      Top             =   525
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   45
      Picture         =   "Controller.frx":8E9D
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Select Area and Starts Capturing"
      Top             =   75
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H006C565C&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2355
      Picture         =   "Controller.frx":9681
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   75
      Width           =   585
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H006C565C&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2965
      Picture         =   "Controller.frx":9AC3
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   75
      Width           =   600
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H006C565C&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4260
      Picture         =   "Controller.frx":A38D
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "AVI to MPG"
      Top             =   450
      Width           =   570
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00000000&
      Caption         =   "step   vbcvbcv"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   810
      Picture         =   "Controller.frx":AC57
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Step Capture"
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton Record 
      Caption         =   "Record"
      Height          =   435
      Left            =   5685
      TabIndex        =   6
      Top             =   3435
      Width           =   1050
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      Height          =   435
      Left            =   6855
      TabIndex        =   5
      Top             =   3450
      Width           =   720
   End
   Begin VB.CommandButton Play 
      Caption         =   "Play"
      Height          =   435
      Left            =   7635
      TabIndex        =   4
      Top             =   3420
      Width           =   720
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000000FF&
      Caption         =   "Õ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Delete Selected Picture"
      Top             =   3210
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6930
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5970
      Pattern         =   "*.bmp"
      TabIndex        =   0
      Top             =   2970
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AVI 2 GIF"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Select AVI to Convert to GIF with Transparency"
      Top             =   1575
      Width           =   705
   End
   Begin VB.PictureBox picstart 
      AutoSize        =   -1  'True
      Height          =   555
      Left            =   -345
      ScaleHeight     =   495
      ScaleWidth      =   435
      TabIndex        =   26
      Top             =   3765
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H006C565C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3585
      Picture         =   "Controller.frx":B521
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ray Mercer AVI 2 BMP"
      Top             =   75
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4365
      Picture         =   "Controller.frx":BDEB
      Top             =   1665
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4080
      Picture         =   "Controller.frx":C6B5
      Top             =   1605
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Capture"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   22
      Top             =   690
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      Height          =   195
      Left            =   4935
      TabIndex        =   18
      Top             =   2895
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Capture to Automate or Select Step or Pan before Capture"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   -105
      TabIndex        =   17
      Top             =   3600
      Width           =   5535
   End
   Begin VB.Image tmpimg 
      Height          =   660
      Left            =   7695
      Top             =   2115
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PauseFlag As Boolean
Dim FagFlag As Boolean
Dim Pan As Boolean, Append As Boolean

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check2.Value = 0 Then
    SetTopMostWindow Me.hWnd, False
Else
    SetTopMostWindow Me.hWnd, True
End If

End Sub

Private Sub Check3_Click()
On Error Resume Next
If Check3.Value = 1 Then
    Label1.Caption = "Drag and Drop to Select an Area to Step-Capture."
    Command16.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    StepFlag = True
    Command18.BackColor = &H8000&
    If Append = False Then Kill App.Path & "\Images\*.BMP"
    PauseFlag = False
    Load frmCapture
    frmCapture.Show
    frmCapture.Timer1.Enabled = False
    Command18.Enabled = True
    Command5.Enabled = True
    Command10.Enabled = True
    Command1.Enabled = False
Else
    Label1.Caption = "Select Capture to Automate or Select Step or Pan before Capture"
    Check3.Value = 0
    StepFlag = False
    Command18.BackColor = &H0&
    Command5_MouseUp 0, 0, 0, 0
    Command18.Enabled = False
    Command16.Enabled = True
    Check3.Enabled = True
    Command5.Enabled = False
    Command10.Enabled = False

End If
End Sub

Private Sub Check4_Click()
On Error Resume Next
If Check4.Value = 1 Then
    Load Panorama
    Panorama.Show
    Pan = True
    Command16.Enabled = False
    Command18.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Command5.Enabled = True
    Command10.Enabled = True
    MsgBox "Click Capture to Select Pan Capture Window!", vbCritical
Else
    Command16.Enabled = True
    Command12.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    Unload Panorama
    Command5.Enabled = False
    Command10.Enabled = False
    Pan = False
End If
End Sub

Private Sub Check5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Check5.Value = 1 Then
    Append = True
Else
    Append = False
End If

End Sub



Private Sub Combo1_Click()
    LoadImage Combo1.Text
   'Command9_Click
End Sub

Public Sub Command1_Click()
On Error Resume Next
EndFlag = False
Command1.Enabled = False
Command16.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Command5.Enabled = True
Command10.Enabled = True
If StepFlag = False Then
  If Pan = False Then
    If Append = False Then Kill App.Path & "\Images\*.BMP"
    PauseFlag = False
    Load frmCapture
    frmCapture.Show
  Else
    If Append = False Then Kill App.Path & "\Images\*.BMP"
    PauseFlag = False
    Load frmCaptPan
    frmCaptPan.Show
  End If
Else
    Label1.Caption = "Use the foot or SPACE bar to Step Capture!"
End If
End Sub

Private Sub Command10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Label2.Caption = "Resume" Then
    frmCapture.Timer1.Enabled = True
    Label2.Caption = "Pause"
    Command10.Picture = Image2.Picture
Else
    frmCapture.Timer1.Enabled = False
    Label2.Caption = "Resume"
    Command10.Picture = Image1.Picture
End If

End Sub

Private Sub Command11_Click()
    Dim i As Integer
    Dim j, k As Long
    On Error Resume Next
    Kill App.Path & "\Images\" & Combo1.Text
    k = Combo1.ListIndex
    Combo1.Clear
    File1.Path = App.Path & "\Images"
    File1.Refresh
    For i = 0 To File1.ListCount - 1
        Combo1.AddItem File1.List(i)
    Next
    Combo1.ListIndex = k
    Combo1.SetFocus

End Sub

Private Sub Command12_Click()
On Error Resume Next
    Dim Mpgest, Mpger, TheMovie As String
    SetTopMostWindow Me.hWnd, False
    'm2apx3g.exe -b PoppingAssCherry1-7.mpg -f -o7 output.avi
    'avi2mpg1 [-options] inputfile.avi [outputfile.mpg]
    Dim RetVal
    TheMovie = App.Path & "\MyAVI.avi"
    Mpger = App.Path & "\MyMPEG1.mpg"
    'Mpgest = App.Path & "\m2apx3g.exe -b " & TheMovie & " -f -o7 " & Mpger & ".avi"
    Mpgest = App.Path & "\avi2mpg1.exe -n -f1 " & TheMovie & " " & Mpger
    RetVal = Shell(Mpgest, 1)

End Sub



Private Sub Command14_Click()
On Error Resume Next
    If Dir(App.Path & "\Giffy.exe") = "" Then
        MsgBox "You need the program Giffy The GIF Animation Builder" & vbCrLf _
        & "1997 WebReady Corp. Good luck finding it ;)"
    Else
        Dim RetVal
        RetVal = Shell(App.Path & "\Giffy.exe Test.gif", 1)
    End If
End Sub

Private Sub Command15_Click()
    Shell App.Path & "\Premiere-Minus.exe MyAVI.avi"
End Sub

Private Sub Command16_Click()
If Pan = False Then Exit Sub
    Label1.Caption = "Select Capture to Automate or Select Step or Pan before Capture"
    Load Panorama
    Panorama.Show
    Pan = True
    Command16.Enabled = False
    Command18.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Command5.Enabled = True
    Command10.Enabled = True
    Load Panorama
    Panorama.Show
    Pan = True
    Command16.Enabled = False
    Command18.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
End Sub

Private Sub Command18_Click()
If StepFlag = False Then Exit Sub
    On Error Resume Next
    frmCapture.Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i As Long
Load fMain
fMain.Show
File1.Refresh
For i = 0 To File1.ListCount - 1
    fMain.lstDIBList.AddItem App.Path & "\Images\" & File1.List(i)
Next
fMain.cmdWriteAVI_Click
Unload fMain
Set fMain = Nothing
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload fMain
Set fMain = Nothing

End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload fMain
Set fMain = Nothing

End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim RetVal
RetVal = Shell(App.Path & "\AVItoGIF.exe MyAVI.avi", 1)

End Sub

Private Sub Command4_Click()
'On Error Resume Next
'Dim RetVal
'RetVal = Shell(App.Path & "\AVICreator6.exe", 1)
    On Error Resume Next
    Kill App.Path & "\Images\*.*"
    Combo1.Clear
    File1.Path = App.Path & "\Images"
    Combo1.Text = "Available Frames to edit"

End Sub

Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim i As Long
Timer1.Enabled = False
Command1.Enabled = True
Command16.Enabled = True
Command18.Enabled = True
Check3.Value = 0
Check4.Value = 0
Check3.Enabled = True
Check4.Enabled = True
Command10.Enabled = False
Command5.Enabled = False
Label1.Caption = "Select Capture to Automate or Select Step or Pan before Capture"
If Pan = False Then
    Check3.Value = 0
    EndFlag = True
    frmCapture.Timer1.Enabled = False
    File1.Path = App.Path & "\Images"
    File1.Refresh
    Combo1.Clear
    Kill App.Path & "\Images\" & File1.List(0)
    For i = 0 To File1.ListCount - 1
        Combo1.AddItem File1.List(i)
    Next
    Combo1.Text = "Scroll Available Frames"
    Combo1.SetFocus
Else
    Pan = False
    Check3.Value = 0
    EndFlag = True
    frmCaptPan.Timer1.Enabled = False
    File1.Path = App.Path & "\Images"
    File1.Refresh
    Combo1.Clear
    Kill App.Path & "\Images\" & File1.List(0)
    For i = 0 To File1.ListCount - 1
        Combo1.AddItem File1.List(i)
    Next
    Combo1.Text = "Scroll Available Frames"
    Combo1.SetFocus
End If
Pan = False: Append = False
Flicker = True
PauseFlag = False
FagFlag = False
Unload frmCapture
Set frmCapture = Nothing
Unload frmCaptPan
Set frmCaptPan = Nothing
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim RetVal
RetVal = Shell(App.Path & "\AVItoGIF.exe", 1)

End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim ngReturnNumber As Long
ngReturnNumber = ShellExecLaunchFile(App.Path & "\Test.gif", "", App.Path)

End Sub

Private Sub Command8_Click()
On Error Resume Next

Dim ngReturnNumber As Long
ngReturnNumber = ShellExecLaunchFile(App.Path & "\MyAVI.avi", "", App.Path)

End Sub

Private Sub Command9_Click()
On Error Resume Next
tmpimg.Refresh
Picture1.Picture = LoadPicture("")
tmpimg.Picture = LoadPicture(App.Path & "\Images\" & Combo1.Text)   'App.Path & "\Images\180705172743.BMP") 'change to your picture path

Dim xImg, yImg As Single, zz As Single
Dim xPic, yPic As Single
xImg = tmpimg.Width
yImg = tmpimg.Height
xPic = Picture1.Width
yPic = Picture1.Height

Dim xRatio, yRatio As Single
xRatio = xImg / xPic
yRatio = yImg / yPic

If xRatio >= yRatio Then
    Picture1.PaintPicture tmpimg.Picture, 0, 0, (tmpimg.Width / xRatio), (tmpimg.Height / xRatio)
Else
    Picture1.PaintPicture tmpimg.Picture, 0, 0, (tmpimg.Width / yRatio), (tmpimg.Height / yRatio)
End If

End Sub

Private Sub Form_Initialize()
On Error Resume Next
    MkDir App.Path & "\Images"
    MkDir App.Path & "\Images1"
    MkDir App.Path & "\SavedBMPs"
    MkDir App.Path & "\Backup"
    If Dir(App.Path & "\AVItoGIF.exe") = "" Then
        MsgBox "You must compile the AVItoGIF.exe file and " & vbCrLf & "place it in the application directory!"
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload frmCapture: Set frmCapture = Nothing
  If KeyCode = vbKeyDelete Then Command11_Click
End Sub
Private Sub LoadImage(filePath As String)
    Dim X As Long
    Dim Y As Long
    On Error Resume Next
    imgMain.Visible = False
    filePath = App.Path & "\Images\" & filePath
    picstart.Picture = LoadPicture(filePath)
    
    ' Establish the ratio of current image s
    '     ize(picstart) verses
    ' the screen size(picMain) and set the i
    '     mage size(imgmain) to
    ' that ratio..
X = picstart.Width
Y = picstart.Height

'x = Controller.ScaleX(x, vbHimetric, vbTwips)
'y = Controller.ScaleY(y, vbHimetric, vbTwips)
    ' First shrink the image so the sides fi
    '     t

    If X > picMain.Width Then
        Y = Y - (X - picMain.Width)
        X = picMain.Width


        'Do Until x = picMain.Width
            'x = x - 1
            'y = y - 1
        'Loop
    End If
    ' if the image is still too tall, shrink
    '     it some more


    If Y > picMain.Height Then
        X = X - (Y - picMain.Height)
        
        Y = picMain.Height

        'Do Until y = picMain.Width
            'x = x - 1
            'y = y - 1
        'Loop
    End If
    
    imgMain.Width = X
    imgMain.Height = Y
    
    ' Center the image(imgmain) in the main
    '     picture box(picmain)
    imgMain.top = (picMain.Height \ 2) - (imgMain.Height \ 2)
    imgMain.left = (picMain.Width \ 2) - (imgMain.Width \ 2)
    
    ' Now copy the image from the start picb
    '     ox(picstart) into the
    ' display image field (imgmain)
    imgMain.Picture = picstart.Picture
    imgMain.Visible = True
End Sub

Private Sub Form_Load()
Dim i As Long
SetTopMostWindow Me.hWnd, True
'Me.Height = 3795
'Me.Width = 4635
Pan = False: Append = False
Me.top = 0      '(Screen.Height - Me.Height) / 2
Me.left = 0     '(Screen.Width - Me.Width) / 2
File1.Path = App.Path & "\Images"
File1.Refresh
For i = 0 To File1.ListCount - 1
    Combo1.AddItem File1.List(i)
Next
Flicker = True
PauseFlag = False
StepFlag = False
EndFlag = False
FagFlag = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Combo1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i
i = mciSendString("close capture", 0&, 0, 0)
Unload frmCapture
Set frmCapture = Nothing
Unload Me
Set Controller = Nothing
End
End Sub



Private Sub Option2_Click()
Flicker = False
End Sub

Private Sub Option2_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload frmCapture: Set frmCapture = Nothing
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Dim i As Integer
    Dim j, k As Long
    On Error Resume Next
    'MsgBox App.Path & "\Images\" & Combo1.Text
    Kill App.Path & "\Images\" & Combo1.Text
    k = Combo1.ListIndex
    Combo1.Clear
    File1.Path = App.Path & "\Images"
    File1.Refresh
    For i = 0 To File1.ListCount - 1
        Combo1.AddItem File1.List(i)
    Next
    Combo1.ListIndex = k
    Combo1.SetFocus
End If

End Sub


Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Combo1.SetFocus

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Combo1.SetFocus

End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Combo1.SetFocus

End Sub




Private Sub Timer1_Timer()
'Add To A Timer With An Interval Of 1
Dim keyresult As Long
keyresult = GetAsyncKeyState(32)
If keyresult = 0 Then Exit Sub
If keyresult = -32767 Then
    Command18_Click
End If


End Sub

