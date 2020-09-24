VERSION 5.00
Begin VB.Form frmCaptPan 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   DrawStyle       =   1  'Dash
   DrawWidth       =   2
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Modern"
      Size            =   14.25
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmCaptPan.frx":0000
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   45
      Top             =   2295
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   90
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   108
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmCaptPan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type PCURSORINFO
    cbSize As Long
    Flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type
'To grab cursor shape -require at least win98 as per Microsoft documentation...
Private Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As PCURSORINFO) As Long
'To get a Handle to the cursor
Private Declare Function GetCursor Lib "user32" () As Long
'To draw cursor shape on bitmap
Private Declare Function DrawIcon Lib "user32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
     
'to get the cursor position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'to end a waiting loopp
Dim GotIt As Boolean
'To use the scrollbars
Dim lngVer As Long
Dim lngHor As Long
Const iconSize As Integer = 9
Dim blnCapturing As Boolean
Const INVERSE = 6       ' DrawMode property - XOR
Const SOLID = 0         ' DrawStyle property
Const DOT = 2           ' DrawStyle property
Dim DrawBox As Boolean
Dim OldX As Single
Dim OldY As Single
Dim StartX As Single
Dim StartY As Single
Dim FirstX As Single, FirstY As Single, SecondX As Single, SecondY As Single
Dim XYZ As Long

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload frmCapture: Set frmCapture = Nothing

End Sub

Private Sub Form_Load()
  'Capture desktop and make it this forms background picture
    Me.DrawStyle = DOT
    'Me.DrawWidth = 2
    'Me.ForeColor = &HFF0000
    DrawBox = True
    Dim DeskhWnd As Long, DeskDC As Long
    Me.WindowState = vbMaximized
    DeskhWnd& = GetDesktopWindow()
    DeskDC& = GetDC(DeskhWnd&)
    BitBlt Me.HDC, 0&, 0&, Screen.Width, Screen.Height, DeskDC&, 0&, 0&, SRCCOPY
    Me.Picture = Me.Image
    UpFlag = False
    blnCapturing = True
    MousePointer = 2
    Timer1.Enabled = True
    Eye = 0
    PicIndex = 0
    'PictBuff.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'User pressed escape so unload
  If KeyCode = vbKeyEscape Then Timer1.Enabled = False: Unload frmCapture: Set frmCapture = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Cls
' Store the initial start of the line to draw.
StartX = x
StartY = y
XX1 = x
YY1 = y
FirstX = x
FirstY = y
' Make the last location equal the starting location
OldX = StartX
OldY = StartY
'Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' If the button is pressed
If Button = 1 Then
' Erase the previous line
    DrawLine StartX, StartY, OldX, OldY
' Draw the new line.
    DrawLine StartX, StartY, x, y
' Save the coordinates for the next call.
    OldX = x
    OldY = y
    SecondX = x
    SecondY = y
    XX2 = x
    YY2 = y
End If
End Sub



Sub DrawLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
' Save the current mode so that you can reset it on
' exit from this sub routine. Not needed in the sample
' but would need it if you are not sure what the
' DrawMode was on entry to this procedure.
Dim SavedMode As Long
SavedMode = Me.DrawMode

' Set to XOR
Me.DrawMode = INVERSE
' Draw a box or line

If DrawBox = True Then
    Me.Line (X1, Y1)-(X2, Y2), , B
Else
    Me.Line (X1, Y1)-(X2, Y2)
End If
'Me.DrawMode = 16
If DrawBox = True Then
    Me.Line (X1 + 2, Y1 + 2)-(X2 + 2, Y2 + 2), , B
Else
    Me.Line (X1 + 2, Y1 + 2)-(X2 + 2, Y2 + 2)
End If

' Reset the DrawMode
Me.DrawMode = SavedMode
End Sub

Public Sub CaptureIt(xStart As Single, xEnd As Single, yStart As Single, yEnd As Single)
  Dim left As Long, top As Long, right As Long, bottom As Long
  Dim lWidth As Long, lHeight As Long
  Dim Eyelet As String
 'On Error Resume Next
  blnCapturing = False
  
  'Get left, right, top and bottom regarldess of where they started and ended
  left = IIf(xStart > xEnd, xEnd, xStart)
  right = IIf(xStart < xEnd, xEnd, xStart)
  top = IIf(yStart > yEnd, yEnd, yStart)
  bottom = IIf(yStart < yEnd, yEnd, yStart)
  lWidth = (right - left)
  lHeight = (bottom - top)
  
  If lWidth <= 0 Or lHeight <= 0 Then GoTo PROC_TOOSMALL  'Nothing to capture
  strOut = lWidth & "x" & lHeight 'Setup our dimensions string
  
  With picTemp
    .Cls  'Clear our picture box that holds the image till copied to clipboar
    .Width = lWidth 'Set it's hight and width
    .Height = lHeight
  End With
  
  Me.Cls  'Clear screen so we don't get the box and dimensions
  BitBlt picTemp.HDC, 0, 0, lWidth, lHeight, Me.HDC, left, top, SRCCOPY 'Copy screen to picture box

clearo:
   'now to get the icon of mouse and paint on form the mouse
   Dim Point As POINTAPI
   GetCursorPos Point
   Dim pcin As PCURSORINFO
   pcin.hCursor = GetCursor
   pcin.cbSize = Len(pcin)
   Dim ret
   ret = GetCursorInfo(pcin)
   'MsgBox iconSize          Point.x & "   " & Point.y
   DrawIcon picTemp.HDC, Point.x - iconSize - left, Point.y - iconSize - top, pcin.hCursor
   'The following paint only mouse shape for this app
   'DrawIcon picTemp.hDC, Point.X - iconSize, Point.Y - iconSize, pcin.hCursor
   'assign to picture the image
   Set picTemp.Picture = picTemp.Image
   'signal you've done to exit the waiting loop


  'Clipboard.Clear 'Clear clipboard
  'Clipboard.SetData picTemp.Image 'Copy image to clipboard

    'SavePicture picTemp.Image, App.Path & "\Images\" & Format(Now, "ddmmyyhhmmss") & ".BMP"
    'PictBuff.Pict(PicIndex).Picture = picTemp.Image
    'PicIndex = PicIndex + 1
    'Load PictBuff.Pict(PicIndex)
    'PictBuff.Caption = PicIndex
    Eyelet = Trim(Str(Eye))
    Do While Len(Trim(Eyelet)) < 5
        Eyelet = "0" & Trim(Eyelet)
    Loop
    Me.WindowState = vbMinimized
    SavePicture picTemp.Image, App.Path & "\Images\" & Eyelet & ".bmp"
    Eye = Eye + 1
    UpFlag = True
    Me.WindowState = vbMinimized
    Me.Visible = False
    XYZ = XYZ + 1
    If XYZ = 1100 Then Timer1.Enabled = False: Unload Me: Exit Sub
    Timer1.Enabled = True
    
    
  'If Clipboard.GetData = 0 Then
    'GoTo clearo
  'Else
    'Load Viewer
    'Viewer.Show
  'End If
PROC_EXIT:
  Exit Sub
  
PROC_TOOSMALL:
  GoTo PROC_EXIT
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'UpFlag = True
'CaptureIt X1, X1a, Y1, Y1a  'Do the capture

    MousePointer = 0
Me.Cls
frmTop.Show
frmBottom.Show
frmLeft.Show
frmRight.Show

frmTop.top = (YY1 - 10) * Screen.TwipsPerPixelY
frmTop.left = (XX1 - 10) * Screen.TwipsPerPixelX
frmTop.Width = (XX2 - XX1 + 20) * Screen.TwipsPerPixelX

frmBottom.top = (YY2 + 10) * Screen.TwipsPerPixelY
frmBottom.left = (XX1 - 10) * Screen.TwipsPerPixelX
frmBottom.Width = (XX2 - XX1 + 20) * Screen.TwipsPerPixelX

frmLeft.top = (YY1 - 10) * Screen.TwipsPerPixelY
frmLeft.left = (XX1 - 10) * Screen.TwipsPerPixelX
frmLeft.Height = (YY2 - YY1 + 20) * Screen.TwipsPerPixelY

frmRight.top = (YY1 - 10) * Screen.TwipsPerPixelY
frmRight.left = (XX2 + 10) * Screen.TwipsPerPixelX
frmRight.Height = (YY2 - YY1 + 26) * Screen.TwipsPerPixelY

Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
Unload Me
Unload frmTop
Set frmTop = Nothing
Unload frmBottom
Set frmBottom = Nothing
Unload frmLeft
Set frmLeft = Nothing
Unload frmRight
Set frmRight = Nothing

Set frmCaptPan = Nothing
'If Controller.Check1.Value = 1 Then
    'Controller.Stop_Click
'End If
End Sub




Private Sub Timer1_Timer()
  
  Timer1.Enabled = False
  Dim DeskhWnd As Long, DeskDC As Long
  DeskhWnd& = GetDesktopWindow()
  DeskDC& = GetDC(DeskhWnd&)
  BitBlt Me.HDC, 0&, 0&, Screen.Width, Screen.Height, DeskDC&, 0&, 0&, SRCCOPY
  Me.Picture = Me.Image
  Me.Visible = True
  'frmCapture.SetFocus
  CaptureIt OldX, StartX, OldY, StartY
  
  OldX = OldX + 1: StartX = StartX + 1 ': OldY = OldY + 1: StartY = StartY + 1
frmTop.top = (YY1 - 10) * Screen.TwipsPerPixelY
frmTop.left = (StartX - 10) * Screen.TwipsPerPixelX
frmTop.Width = (XX2 - XX1 + 20) * Screen.TwipsPerPixelX

frmBottom.top = (YY2 + 10) * Screen.TwipsPerPixelY
frmBottom.left = (StartX - 10) * Screen.TwipsPerPixelX
frmBottom.Width = (XX2 - XX1 + 20) * Screen.TwipsPerPixelX

frmLeft.top = (YY1 - 10) * Screen.TwipsPerPixelY
frmLeft.left = (StartX - 10) * Screen.TwipsPerPixelX
frmLeft.Height = (YY2 - YY1 + 20) * Screen.TwipsPerPixelY

frmRight.top = (YY1 - 10) * Screen.TwipsPerPixelY
frmRight.left = (StartX + (XX2 - XX1 + 10)) * Screen.TwipsPerPixelX
frmRight.Height = (YY2 - YY1 + 26) * Screen.TwipsPerPixelY

End Sub
