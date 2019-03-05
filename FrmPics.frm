VERSION 5.00
Begin VB.Form FrmPics 
   Caption         =   "Form2"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   0
      Picture         =   "FrmPics.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   11640
      Left            =   -120
      ScaleHeight     =   772
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1028
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15480
   End
End
Attribute VB_Name = "FrmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
