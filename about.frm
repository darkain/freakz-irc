VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "About DooVoo Chat"
   ClientHeight    =   3840
   ClientLeft      =   1545
   ClientTop       =   1095
   ClientWidth     =   7275
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   6120
      TabIndex        =   0
      Top             =   3240
      Width           =   976
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "DooVoo Chat Copyright (C) 2000 by Vincent E. Milum Jr."
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   3960
      TabIndex        =   1
      Top             =   3240
      Width           =   2040
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3870
      Left            =   0
      Picture         =   "about.frx":030A
      Top             =   0
      Width           =   7290
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XX As Integer, YY As Integer, ZZ As Byte
Dim CurX As Integer, CurY As Integer

Private Sub Form_Load()
Me.Left = (FrmMain.Width \ 2) - (Me.Width \ 2)
Me.Top = (FrmMain.Height \ 2) - (Me.Height \ 2)
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
XX = X
YY = Y
ZZ = 1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ZZ = 1 Then
  Left = Left + X - XX
  Top = Top + Y - YY
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ZZ = 0
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
XX = X
YY = Y
ZZ = 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ZZ = 1 Then
  Left = Left + X - XX
  Top = Top + Y - YY
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ZZ = 0
End Sub

Private Sub OK_Click()
Unload Me
End Sub
