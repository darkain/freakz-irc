VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FrmDebug 
   Caption         =   "Debug Window"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmDebug.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox Incoming 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      _Version        =   327680
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmDebug.frx":030A
   End
End
Attribute VB_Name = "FrmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AddText(textmsg As String)
Incoming.SelBold = False
Incoming.SelItalic = False
Incoming.SelUnderline = False
Incoming.SelFontName = "MS Sans Serif"
Incoming.SelStart = Len(Incoming.Text)
Incoming.SelText = CRLF
Incoming.SelColor = RGB(255, 0, 0)
Incoming.SelText = "["
Incoming.SelColor = RGB(0, 127, 255)
Incoming.SelText = Time
Incoming.SelColor = RGB(255, 0, 0)
Incoming.SelText = "] "
Incoming.SelColor = RGB(0, 255, 255)
Incoming.SelText = textmsg
End Sub

Private Sub Form_Load()
Incoming.Text = ""
Incoming.SelColor = RGB(0, 128, 255)
End Sub

Private Sub Form_Resize()
On Error Resume Next
Incoming.Width = ScaleWidth
Incoming.Height = ScaleHeight
End Sub
