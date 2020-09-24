VERSION 5.00
Begin VB.Form Panorama 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9540
   Icon            =   "Panorama.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Panorama.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   7200
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9300
      Top             =   210
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   345
      Left            =   18150
      TabIndex        =   1
      Top             =   6870
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   7260
      Left            =   -90
      Picture         =   "Panorama.frx":1194
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   -60
      Width           =   9660
      Begin Capture2AVIGIF.AnyShape AnyShape1 
         Height          =   1635
         Left            =   6135
         TabIndex        =   2
         Top             =   1470
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   2884
         Picture         =   "Panorama.frx":2839D
         MaskColor       =   14999268
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Image tmpimg 
      Height          =   3165
      Left            =   255
      Top             =   180
      Visible         =   0   'False
      Width           =   3945
   End
End
Attribute VB_Name = "Panorama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xRatio, yRatio As Single
Dim xImg, yImg As Single
Dim xPic, yPic As Single

Private Sub Form_Activate()
Picture1.top = 0
Picture1.left = 0
Me.Width = Picture1.Width
Me.Height = Picture1.Height

End Sub

Private Sub Form_Load()
Me.top = 0
Me.left = 0
Picture1.top = 0
Picture1.left = 0
Me.Width = Picture1.Width
Me.Height = Picture1.Height

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
AnyShape1.left = x
AnyShape1.top = y - 2000
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
'    If Not IsWindow(hToolbar) Then
'        MsgBox "This sample only works when run as an executable. Please see the comments section at the top of this form's source code."
'        Exit Sub
'    End If
    Dim R As RECT, P As POINTAPI
    GetWindowRect hToolbar, R
    P.x = R.right - 38
    P.y = R.top + 10
    SetCursorPos P.x, P.y
    mouse_event MOUSEEVENTF_LEFTDOWN, P.x, P.y, 0, 0
    SetCursorPos P.x, P.y + 24
    mouse_event MOUSEEVENTF_LEFTUP, P.x, P.y, 0, 0
    mouse_event MOUSEEVENTF_LEFTDOWN, P.x, P.y, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, P.x, P.y, 0, 0
    hToolbar = 0
End Sub
