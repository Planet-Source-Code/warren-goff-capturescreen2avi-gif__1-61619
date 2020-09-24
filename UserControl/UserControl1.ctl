VERSION 5.00
Begin VB.UserControl AnyShape 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   990
   HitBehavior     =   0  'None
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   495
   ScaleWidth      =   990
   Begin VB.PictureBox pic2 
      BackColor       =   &H00E0E0E0&
      Height          =   2535
      Left            =   5220
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   2535
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00E0E0E0&
      Height          =   2445
      Left            =   2250
      ScaleHeight     =   2385
      ScaleWidth      =   2745
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2805
   End
End
Attribute VB_Name = "AnyShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 

'   CONTROL IS DRAWN OR DBL CLICKED ON FORM FOR FIRST TIME
'          -UserControl_Initialize
'          -UserControl_InitProperties
'          -UserControl_Resize
'          -UserControl_Show
'          -UserControl_Paint
'
'   CONTROL IS RESIZED ON FORM AT DESIGN TIME
'          -UserControl_Resize
'          -UserControl_Paint
'
'   THE "RUN" OR "START" BUTTON IS CLICKED PUTTING CONTROL IN RUN MODE
'          -UserControl_Hide
'          -UserControl_Terminate
'          -UserControl_Initialize
'          -UserControl_Resize
'          -UserControl_ReadProperties
'          -UserControl_Show
'          -UserControl_EnterFocus
'          -UserControl_GotFocus
'          -UserControl_Paint
'
'    THE FORM IS CLOSED OR TEMINATED, PUTTING CONTROL BACK IN DESIGN MODE
'          -UserControl_LostFocus
'          -UserControl_ExitFocus
'          -UserControl_Hide
'          -UserControl_Terminate
'          -UserControl_Initialize
'          -UserControl_Resize
'          -UserControl_ReadProperties
'          -UserControl_Show
'          -UserControl_Paint
'
'    YOU CHANGE YOUR PROJECT FROM CODE VIEW TO DESIGN (FORM) VIEW
'          -UserControl_Paint
'
'    THE CONTROL IS REMOVED FROM THE FORM
'          -UserControl_WriteProperties
'          -UserControl_Hide
'          -UserControl_Terminate
'=======================================================================

 

'[EVENTS]
Event Click()
Event MouseEnter()
Event MouseExit()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'[CONSTANTS TOOLTIP]
Private Const COLOR_INFOBK = 24   'default backcolor for tooltip
Private Const COLOR_INFOTEXT = 23 'default text color for tooltip
Private Const WM_USER              As Integer = &H400
Private Const TTF_CENTERTIP        As Integer = &H2
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTM_ADDTOOLA         As Integer = (WM_USER + 4)
Private Const TTM_UPDATETIPTEXTA   As Integer = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH   As Integer = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR    As Integer = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR  As Integer = (WM_USER + 20)
Private Const TTS_ALWAYSTIP        As Integer = &H1
Private Const TTF_SUBCLASS         As Integer = &H10
Private Const CW_USEDEFAULT        As Long = &H80000000

'[CONSTANTS OTHER]
Private Const SRCCOPY = &HCC0020
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const CLR_INVALID = -1

'[TYPES]
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect   As RECT
    hInstance As Long
    lpStr As String
    lParam  As Long
End Type

'[ENUMS]
Private Enum enBS
   enNull = 0
   enEntered = 1
   enDown = 2
End Enum

Public Enum enAppType
    InternetAddress = 0
    LocalFolderFile = 1
End Enum

Public Enum enCaptStyle
   enRegular = 0
   enRaised = 1
   enEmbossed = 2
End Enum

'[API FOR CREATING AND USING TOOLTIP]
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'[API OTHER]
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
 
Dim BtnState As enBS
 

'Default Property Values:
Const m_def_Caption = "Caption"
Const m_def_MultilineToolTipString = ""
Const m_def_HasRestingFrame = False
Const m_def_ImageX = 5
Const m_def_ImageY = 1
Const m_def_DisplayFrame = 0
Const m_def_MouseEnterSound = ""
Const m_def_CaptionStyle = 0
Const m_def_CaptionY = 20
Const m_def_CaptionX = 1
Const m_def_MouseOverCaptionColor = &HFF0000
Const m_def_CaptionColor = 0

'Property Variables:
Dim m_Caption As String
Dim m_MultilineToolTipString As String
Dim m_ToolTipBackcolor As OLE_COLOR
Dim m_ToolTipForecolor As OLE_COLOR
Dim m_HasRestingFrame As Boolean
Dim m_ImageX As Long
Dim m_ImageY As Long
Dim m_DisplayFrame As Boolean
Dim m_MouseEnterSound As String
Dim m_CaptionStyle As enCaptStyle
Dim m_CaptionY As Long
Dim m_CaptionX As Long
Dim m_MouseOverCaptionColor As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR

'[private variables]
Private lHwnd                      As Long
Private ti                         As TOOLINFO
Private mvarTipText                As String


 
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
   Debug.Print "UserControl_AccessKeyPress"
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
   Debug.Print "UserControl_AmbientChanged"
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    Debug.Print "UserControl_AsyncReadComplete"
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    Debug.Print "UserControl_AsyncReadProgress"
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub
 

 

Private Sub UserControl_EnterFocus()
   Debug.Print "UserControl_EnterFocus"
End Sub

Private Sub UserControl_ExitFocus()
   Debug.Print "UserControl_ExitFocus"
End Sub

Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
    Debug.Print "UserControl_GetDataMember"
End Sub

Private Sub UserControl_GotFocus()
    Debug.Print "UserControl_GotFocus"
End Sub

Private Sub UserControl_Hide()
   Debug.Print "UserControl_Hide"
End Sub

Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    Debug.Print "UserControl_HitTest"
End Sub

Private Sub UserControl_Initialize()
    Debug.Print "UserControl_Initialize"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "UserControl_KeyDown"
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   Debug.Print "UserControl_KeyPress"
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "UserControl_KeyUp"
End Sub

Private Sub UserControl_LostFocus()
   Debug.Print "UserControl_LostFocus"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BtnState = enDown
    Call UserControl_Paint
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------
' create the mouseenter
' and exits
'---------------------
     If (x < 0) Or (y < 0) Or (x > ScaleWidth) Or (y > ScaleHeight) Then
              ReleaseCapture
              BtnState = enNull
              Call UserControl_Paint
              'mouseexit event
              RaiseEvent MouseExit
                           
     ElseIf GetCapture() <> hWnd Then
              SetCapture hWnd
              BtnState = enEntered
              Call UserControl_Paint
              'mouseenter event
              RaiseEvent MouseEnter
              'play sound if specified
              If m_MouseEnterSound <> "" Then
                  PlaySound m_MouseEnterSound, 0&, SND_ASYNC Or SND_NODEFAULT
              End If
     End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BtnState = enNull
    Call UserControl_Paint
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
'
  Dim plusBuff&

   'clear previous image
   pic2.Cls
    
   'if the mouse is down, we paint the image
   '1 pixel further to the right and down to
   'enhance a mousedown effect
   If BtnState = enDown Then
       plusBuff = 1
   Else
       plusBuff = 0
   End If
   
   'paint the buttons image, pic1, to pic2
   If pic1.Picture <> 0 Then BitBlt pic2.HDC, _
         (m_ImageX + plusBuff), (m_ImageY + plusBuff), _
          ScaleWidth, ScaleHeight, pic1.HDC, 0, 0, SRCCOPY
   
   'If displayframe= true then paint a button
   'raised effect when mouse is over, and button
   'depressed effect when mouse is down
   If m_DisplayFrame = True Then
      If BtnState = enEntered Then
          Call PaintBorder
      ElseIf BtnState = enDown Then
          Call PaintBorder(True)
      End If
   End If
   
   'print the caption to pic2m if there is one
   If m_Caption <> "" Then _
          Call PrintCaption(Val(BtnState))
   
   'paint the image and the caption
   'now on pic2, to the usercontrol
   TransparentBlt HDC, 0, 0, ScaleWidth, ScaleHeight, pic2.HDC, _
                 0, 0, ScaleWidth, ScaleHeight, RGB(255, 0, 255)
   
   'set the visible part or mask of this usercontrol
   'to the image in picture2, which is what we just painted
   UserControl.BackStyle = 0
   UserControl.MaskPicture = pic2.Image

End Sub


Private Sub PaintBorder(Optional MouseIsDown As Boolean = False)
'
Dim clr1&, clr2&
'
'paints a 3d button edges effect when mouse is down and when
'mouse has entered the control
'
If MouseIsDown Then
   clr2 = vbWhite
   clr1 = RGB(130, 130, 130)
Else
   clr2 = RGB(130, 130, 130)
   clr1 = vbWhite
End If
'
'this line is painted at the left edge and top edge
pic2.Line (1, 1)-(ScaleWidth, ScaleHeight), clr1, B
'this line is painted at the right and bottom edge
pic2.Line (-1, -1)-((ScaleWidth - 2), (ScaleHeight - 2)), clr2, B
'
End Sub

Private Sub PrintCaption(Optional bMouseEnter As Boolean = False)
'----------------------
Dim txtWid&, txtHei&
Dim lPos&, tPos&

'print the caption in white just to right and below or
'just to left and above if caption style is embossed or raised
If m_CaptionStyle = enRaised Then
   SetTextColor pic2.HDC, vbWhite
   TextOut pic2.HDC, (m_CaptionX - 1), (m_CaptionY - 1), m_Caption, Len(m_Caption)
       
ElseIf m_CaptionStyle = enEmbossed Then
    SetTextColor pic2.HDC, vbWhite
    TextOut pic2.HDC, (m_CaptionX + 2), (m_CaptionY + 1), m_Caption, Len(m_Caption)
        
End If
     
'print caption the right color based on whether
'mouse is over control or not
If bMouseEnter = True Then
     SetTextColor pic2.HDC, m_MouseOverCaptionColor
Else
     SetTextColor pic2.HDC, m_CaptionColor
End If
             
'prints the text
TextOut pic2.HDC, m_CaptionX, m_CaptionY, m_Caption, Len(m_Caption)
'
End Sub


Private Sub UserControl_Resize()
  Debug.Print "UserControl_Resize"
  'enforce min and max width and height
   If Width < 300 Then
      Width = 300
   ElseIf Width > 2000 Then
      Width = 2000
   End If
   If Height < 300 Then
      Height = 300
   ElseIf Height > 2000 Then
      Height = 2000
   End If
End Sub

Private Sub UserControl_Show()
   '
   'initialize the constituate controls for proper painting
   '
   With pic1
      .ScaleMode = 3 'pixels
      .AutoRedraw = True
      .AutoSize = False
      .BorderStyle = 0 'none
      .Appearance = 0 'flat
      .BackColor = UserControl.MaskColor
   End With
   With pic2
      .ScaleMode = 3 'pixels
      .AutoRedraw = True
      .AutoSize = False
      .BorderStyle = 0 'none
      .Appearance = 0 'flat
      .BackColor = UserControl.MaskColor
   End With
   With UserControl
      .ScaleMode = 3 'pixels
      .Appearance = 0 'flat
      .BackStyle = 1
      .AutoRedraw = False 'were doing the painting
   End With
   '
End Sub
 








'===========================================================
' FOR CONVERT A LONG COLOR(OLE COLOR) TO RGB
'===========================================================
Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal& = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then _
            TranslateColor = CLR_INVALID
End Function
Private Sub LongToRGB(ByVal lngColor As Long, intRed As Integer, intGreen As Integer, intBlue As Integer)
    Dim LRGB&

    LRGB = TranslateColor(lngColor)
    
    If LRGB <> -1 Then
      intRed = LRGB Mod &H100
      LRGB = (LRGB \ &H100)
      intGreen = LRGB Mod &H100
      LRGB = (LRGB \ &H100)
      intBlue = LRGB Mod &H100
    End If
End Sub
'============================================================


'===========================================================
' CREATE THE TOOLTIP
'===========================================================
Private Function CreateTooltip() As Boolean
  Dim lpRect As RECT
  Dim R%, G%, B%
  Exit Function
        'if there is already a previous tooltip here destroy it
        If lHwnd <> 0 Then DestroyWindow lHwnd
 
          'create the tooltip
         lHwnd = CreateWindowEx( _
                0&, "tooltips_class32", vbNullString, TTS_ALWAYSTIP Or TTS_NOPREFIX, _
                CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
                hWnd, 0&, App.hInstance, 0& _
                )
         'tooltip on top
         Call SetWindowPos(lHwnd, -1, 0&, 0&, 0&, 0&, &H1 Or &H2 Or &H10)
         
        'get and save the rect coodinates for the ctl with the tooltip
         GetClientRect hWnd, lpRect
        
        With ti
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            .lHwnd = hWnd
            .lId = 0
            .hInstance = App.hInstance
            .lpRect = lpRect
        End With
        
        SendMessage lHwnd, TTM_ADDTOOLA, 0&, ti
        
        'this one allows the multiline effect
        SendMessage lHwnd, TTM_SETMAXTIPWIDTH, 200, 0
        
        'tooltip text color
        Call LongToRGB(m_ToolTipForecolor, R, G, B)
        SendMessage lHwnd, TTM_SETTIPTEXTCOLOR, RGB(R, G, B), 0&
        
        'tooltip backcolor
        Call LongToRGB(m_ToolTipBackcolor, R, G, B)
        SendMessage lHwnd, TTM_SETTIPBKCOLOR, RGB(R, G, B), 0&
   
End Function
'===========================================================
 






Public Sub LaunchApp(AppType As enAppType, spath$)
Attribute LaunchApp.VB_Description = "Will open a file or folder or launch an internet address with your default browser, normally called from this controls click or mousedown event"
'-------------------------------------------
' launch file or internet address
'-------------------------------------------
   Call ShellExecute( _
            hWnd, "open", spath, vbNullString, vbNullString, 1 _
            )
End Sub

 

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   Dim lng&
   Dim ttTextR%, ttTextG%, ttTextB%
   Dim ttBackR%, ttBackG%, ttBackB%
      
        'get the syst def tooltip text color, convert to rgb
        lng& = GetSysColor(COLOR_INFOTEXT)
        Call LongToRGB(lng&, ttTextR, ttTextG, ttTextB)
        
        'do same for sys def tooltip back color
        lng& = GetSysColor(COLOR_INFOBK)
        Call LongToRGB(lng&, ttBackR, ttBackG, ttBackB)
        
        m_Caption = m_def_Caption
        m_CaptionColor = m_def_CaptionColor
        m_MouseOverCaptionColor = m_def_MouseOverCaptionColor
        m_CaptionX = m_def_CaptionX
        m_CaptionY = m_def_CaptionY
        m_CaptionStyle = m_def_CaptionStyle
        m_MouseEnterSound = m_def_MouseEnterSound
        m_DisplayFrame = m_def_DisplayFrame
        m_ImageX = m_def_ImageX
        m_ImageY = m_def_ImageY
        m_HasRestingFrame = m_def_HasRestingFrame
        m_ToolTipForecolor = RGB(ttTextR, ttTextG, ttTextB)
        m_ToolTipBackcolor = RGB(ttBackR, ttBackG, ttBackB)
        m_MultilineToolTipString = m_def_MultilineToolTipString
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Dim lng&
   Dim ttTextR%, ttTextG%, ttTextB%
   Dim ttBackR%, ttBackG%, ttBackB%
        
        Debug.Print "UserControl_ReadProperties"
        'get the syst def tooltip text color, convert to rgb
        lng& = GetSysColor(COLOR_INFOTEXT)
        Call LongToRGB(lng&, ttTextR, ttTextG, ttTextB)
        
        'do same for sys def tooltip back color
        lng& = GetSysColor(COLOR_INFOBK)
        Call LongToRGB(lng&, ttBackR, ttBackG, ttBackB)
        
        
        Set pic1.Picture = PropBag.ReadProperty("Picture", Nothing)
        UserControl.MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
        m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
        m_CaptionColor = PropBag.ReadProperty("CaptionColor", m_def_CaptionColor)
        m_MouseOverCaptionColor = PropBag.ReadProperty("MouseOverCaptionColor", m_def_MouseOverCaptionColor)
        m_CaptionX = PropBag.ReadProperty("CaptionX", m_def_CaptionX)
        m_CaptionY = PropBag.ReadProperty("CaptionY", m_def_CaptionY)
        m_CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
        m_MouseEnterSound = PropBag.ReadProperty("MouseEnterSound", m_def_MouseEnterSound)
        m_DisplayFrame = PropBag.ReadProperty("DisplayFrame", m_def_DisplayFrame)
        m_ImageX = PropBag.ReadProperty("ImageX", m_def_ImageX)
        m_ImageY = PropBag.ReadProperty("ImageY", m_def_ImageY)
        m_HasRestingFrame = PropBag.ReadProperty("HasRestingFrame", m_def_HasRestingFrame)
        m_ToolTipForecolor = PropBag.ReadProperty("ToolTipForecolor", RGB(ttTextR, ttTextG, ttTextB))
        m_ToolTipBackcolor = PropBag.ReadProperty("ToolTipBackcolor", RGB(ttBackR, ttBackG, ttBackB))
        m_MultilineToolTipString = PropBag.ReadProperty("MultilineToolTipString", m_def_MultilineToolTipString)

        If m_HasRestingFrame = True Then
           UserControl.BorderStyle = 1
        Else
           UserControl.BorderStyle = 0 'none
        End If
        
        If Ambient.UserMode = False Then
            'destroy the tooltip window
            If lHwnd <> 0 Then DestroyWindow lHwnd
        Else
            'create tooltip
            Call CreateTooltip
            ti.lpStr = m_MultilineToolTipString
            If lHwnd <> 0 Then SendMessage lHwnd, TTM_UPDATETIPTEXTA, 0&, ti
        End If
        Set pic2.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub
 
Private Sub UserControl_Terminate()
    Debug.Print "UserControl_Terminate"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Dim lng&
   Dim ttTextR%, ttTextG%, ttTextB%
   Dim ttBackR%, ttBackG%, ttBackB%
         
         Debug.Print "UserControl_WriteProperties"
        'get the syst def tooltip text color, convert to rgb
        lng& = GetSysColor(COLOR_INFOTEXT)
        Call LongToRGB(lng&, ttTextR, ttTextG, ttTextB)
        
        'do same for sys def tooltip back color
        lng& = GetSysColor(COLOR_INFOBK)
        Call LongToRGB(lng&, ttBackR, ttBackG, ttBackB)
        
        
        Call PropBag.WriteProperty("Picture", pic1.Picture, Nothing)
        Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, &HFF00FF)
        Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
        Call PropBag.WriteProperty("CaptionColor", m_CaptionColor, m_def_CaptionColor)
        Call PropBag.WriteProperty("MouseOverCaptionColor", m_MouseOverCaptionColor, m_def_MouseOverCaptionColor)
        Call PropBag.WriteProperty("CaptionX", m_CaptionX, m_def_CaptionX)
        Call PropBag.WriteProperty("CaptionY", m_CaptionY, m_def_CaptionY)
        Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
        Call PropBag.WriteProperty("MouseEnterSound", m_MouseEnterSound, m_def_MouseEnterSound)
        Call PropBag.WriteProperty("DisplayFrame", m_DisplayFrame, m_def_DisplayFrame)
        Call PropBag.WriteProperty("ImageX", m_ImageX, m_def_ImageX)
        Call PropBag.WriteProperty("ImageY", m_ImageY, m_def_ImageY)
        Call PropBag.WriteProperty("HasRestingFrame", m_HasRestingFrame, m_def_HasRestingFrame)
        Call PropBag.WriteProperty("ToolTipForecolor", m_ToolTipForecolor, RGB(ttTextR, ttTextG, ttTextB))
        Call PropBag.WriteProperty("ToolTipBackcolor", m_ToolTipBackcolor, RGB(ttBackR, ttBackG, ttBackB))
        Call PropBag.WriteProperty("MultilineToolTipString", m_MultilineToolTipString, m_def_MultilineToolTipString)
        Call PropBag.WriteProperty("Font", pic2.Font, Ambient.Font)
End Sub
 
'[PICTURE]
Public Property Get Picture() As Picture
        Set Picture = pic1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
        Set pic1.Picture = New_Picture
        PropertyChanged "Picture"
        Call UserControl_Paint
End Property

'[MASKCOLOR]
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
        MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
        UserControl.MaskColor() = New_MaskColor
        PropertyChanged "MaskColor"
        pic1.BackColor = UserControl.MaskColor
        pic2.BackColor = UserControl.MaskColor
        Call UserControl_Paint
End Property

'[CAPTION]
Public Property Get Caption() As String
        Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
        m_Caption = New_Caption
        PropertyChanged "Caption"
        Call UserControl_Paint
End Property

'[CAPTIONCOLOR]
Public Property Get CaptionColor() As OLE_COLOR
        CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
        m_CaptionColor = New_CaptionColor
        PropertyChanged "CaptionColor"
        Call UserControl_Paint
End Property

'[MOUSEOVERCAPTIONCOLOR]
Public Property Get MouseOverCaptionColor() As OLE_COLOR
        MouseOverCaptionColor = m_MouseOverCaptionColor
End Property

Public Property Let MouseOverCaptionColor(ByVal New_MouseOverCaptionColor As OLE_COLOR)
        m_MouseOverCaptionColor = New_MouseOverCaptionColor
        PropertyChanged "MouseOverCaptionColor"
End Property

'[CAPTIONX]
Public Property Get CaptionX() As Long
Attribute CaptionX.VB_Description = "the x position, in pixels of the start of captioin printing"
        CaptionX = m_CaptionX
End Property

Public Property Let CaptionX(ByVal New_CaptionX As Long)
        m_CaptionX = New_CaptionX
        PropertyChanged "CaptionX"
        Call UserControl_Paint
End Property

'[CAPTIONY]
Public Property Get CaptionY() As Long
Attribute CaptionY.VB_Description = "the y position, in pixels of the start of captioin printing"
        CaptionY = m_CaptionY
End Property

Public Property Let CaptionY(ByVal New_CaptionY As Long)
        m_CaptionY = New_CaptionY
        PropertyChanged "CaptionY"
       
        Call UserControl_Paint
End Property
 
'[IMAGEX]
Public Property Get ImageX() As Long
        ImageX = m_ImageX
End Property

Public Property Let ImageX(ByVal New_ImageX As Long)
        m_ImageX = New_ImageX
        PropertyChanged "ImageX"
        Call UserControl_Paint
End Property

'[IMAGEY]
Public Property Get ImageY() As Long
        ImageY = m_ImageY
End Property

Public Property Let ImageY(ByVal New_ImageY As Long)
        m_ImageY = New_ImageY
        PropertyChanged "ImageY"
        Call UserControl_Paint
End Property
'[CAPTIONSTYLE]
Public Property Get CaptionStyle() As enCaptStyle
Attribute CaptionStyle.VB_Description = "The intensity of shadowing and how dramatically the image moves on mousedown"
        CaptionStyle = m_CaptionStyle
End Property

Public Property Let CaptionStyle(ByVal New_CaptionStyle As enCaptStyle)
        m_CaptionStyle = New_CaptionStyle
        PropertyChanged "CaptionStyle"
        Call UserControl_Paint
End Property


'[MOUSEENTERSOUND]
Public Property Get MouseEnterSound() As String
Attribute MouseEnterSound.VB_Description = "Path to sound file to play when mouse enters the control"
        MouseEnterSound = m_MouseEnterSound
End Property

Public Property Let MouseEnterSound(ByVal New_MouseEnterSound As String)
        m_MouseEnterSound = New_MouseEnterSound
        PropertyChanged "MouseEnterSound"
End Property

'[FONT]
Public Property Get Font() As Font
        Set Font = pic2.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
        Set pic2.Font = New_Font
        PropertyChanged "Font"
        Call UserControl_Paint
End Property

'[DISPLAYFRAME]
Public Property Get DisplayFrame() As Boolean
Attribute DisplayFrame.VB_Description = "Button shows Border highlite and shadows on mouseover and mousedown"
        DisplayFrame = m_DisplayFrame
End Property

Public Property Let DisplayFrame(ByVal New_DisplayFrame As Boolean)
        m_DisplayFrame = New_DisplayFrame
        PropertyChanged "DisplayFrame"
End Property

 
'[HASRESTINGFRAME]
Public Property Get HasRestingFrame() As Boolean
Attribute HasRestingFrame.VB_Description = "I thin frame is drawn around the control"
        HasRestingFrame = m_HasRestingFrame
End Property

Public Property Let HasRestingFrame(ByVal New_HasRestingFrame As Boolean)
        m_HasRestingFrame = New_HasRestingFrame
        PropertyChanged "HasRestingFrame"
        
        If New_HasRestingFrame = True Then
           UserControl.BorderStyle = 1 'none
        Else
           UserControl.BorderStyle = 0
        End If
End Property

'[TOOLTIPFORECOLOR]
Public Property Get ToolTipForecolor() As OLE_COLOR
Attribute ToolTipForecolor.VB_Description = "The color of the text for the controls tooltip"
        ToolTipForecolor = m_ToolTipForecolor
End Property

Public Property Let ToolTipForecolor(ByVal New_ToolTipForecolor As OLE_COLOR)
        m_ToolTipForecolor = New_ToolTipForecolor
        PropertyChanged "ToolTipForecolor"
End Property

'[TOOLTIPBACKCOLOR]
Public Property Get ToolTipBackcolor() As OLE_COLOR
Attribute ToolTipBackcolor.VB_Description = "The backcolor of the controls tooltip"
        ToolTipBackcolor = m_ToolTipBackcolor
End Property

Public Property Let ToolTipBackcolor(ByVal New_ToolTipBackcolor As OLE_COLOR)
        m_ToolTipBackcolor = New_ToolTipBackcolor
        PropertyChanged "ToolTipBackcolor"
End Property
 

'[MULTILINETOOLTIPSTRING]
Public Property Get MultilineToolTipString() As String
Attribute MultilineToolTipString.VB_Description = "The single or multiline tooltip that will display when mouse rests over the control. In order for the tooltip to be multiline, you must set the property in the code window, not this property window"
        MultilineToolTipString = m_MultilineToolTipString
End Property

Public Property Let MultilineToolTipString(ByVal New_MultilineToolTipString As String)
        m_MultilineToolTipString = New_MultilineToolTipString
        PropertyChanged "MultilineToolTipString"
        ti.lpStr = New_MultilineToolTipString
        If lHwnd <> 0 Then SendMessage lHwnd, TTM_UPDATETIPTEXTA, 0&, ti
End Property

 
 

