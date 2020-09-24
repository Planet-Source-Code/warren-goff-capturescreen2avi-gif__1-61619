VERSION 5.00
Begin VB.Form frmLeft 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   LinkTopic       =   "Form2"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   6
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SetTopMostWindow Me.hWnd, True
End Sub
