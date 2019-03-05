VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DooVoo IRC Setup"
   ClientHeight    =   4215
   ClientLeft      =   1560
   ClientTop       =   1410
   ClientWidth     =   5790
   Icon            =   "setup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4215
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog cmd 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Server"
            Key             =   "Server"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Server"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Font"
            Key             =   "Font"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Font"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc"
            Object.Tag             =   "Misc"
            Object.ToolTipText     =   "Misc"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Top             =   3840
      Width           =   1040
   End
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   3840
      Width           =   1040
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      Begin VB.ComboBox ServerCombo 
         Height          =   315
         ItemData        =   "setup.frx":030A
         Left            =   1200
         List            =   "setup.frx":0317
         TabIndex        =   3
         Text            =   "192.168.0.1"
         Top             =   240
         Width           =   1808
      End
      Begin VB.TextBox PortText 
         Height          =   304
         Left            =   1200
         TabIndex        =   2
         Text            =   "6667"
         Top             =   675
         Width           =   1808
      End
      Begin VB.TextBox NickText 
         Height          =   304
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   1
         Text            =   "Darkain_X"
         Top             =   1125
         Width           =   1808
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   750
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nickname:"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "setup.frx":0340
         Left            =   1440
         List            =   "setup.frx":037B
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command1"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   2400
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim Colourz As String
cmd.CancelError = True
On Error GoTo ErrHandle
cmd.Flags = &H3
cmd.Color = Label4.BackColor
cmd.ShowColor
Label4.BackColor = cmd.Color

ErrHandle:
Exit Sub
End Sub

Private Sub Form_Load()
'Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
'Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
ServerCombo.Text = ServerAddress
PortText.Text = Port
NickText = Nickname
Combo1.ListIndex = 0
End Sub

Private Sub OK_Click()

  ' Make sure all fields have data
  If ServerCombo.Text = "" Then
    Beep
    ServerCombo.SetFocus
    Exit Sub
  End If
  If PortText.Text = "" Then
    Beep
    PortText.SetFocus
    Exit Sub
  End If
  If NickText.Text = "" Then
    Beep
    NickText.SetFocus
    Exit Sub
  End If

  ' Set the global variables
  ServerAddress = ServerCombo.Text
  Port = PortText.Text
  Nickname = NickText.Text
  ' Close setup
  Unload Me

End Sub


