VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FrmPrivMsg 
   Caption         =   "Private Message"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "FrmPrivMsg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   6855
   Visible         =   0   'False
   Begin VB.TextBox Outgoing 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1970
      TabIndex        =   1
      Top             =   3360
      Width           =   4035
   End
   Begin ComctlLib.TreeView Tree 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6376
      _Version        =   327680
      Style           =   7
      Appearance      =   1
      MouseIcon       =   "FrmPrivMsg.frx":030A
   End
   Begin RichTextLib.RichTextBox Incoming 
      Height          =   3015
      Left            =   1970
      TabIndex        =   2
      Top             =   0
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmPrivMsg.frx":0326
   End
   Begin RichTextLib.RichTextBox TmpTxt 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   327680
      Enabled         =   -1  'True
      TextRTF         =   $"FrmPrivMsg.frx":03EF
   End
End
Attribute VB_Name = "FrmPrivMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastMessage
Dim PauseOutput As Byte

Sub SendData(textmsg As String)
FrmMain.SendData textmsg
End Sub

Sub AddText(textmsg As String)
If PauseOutput = False Then
  TmpTxt.Text = ""
  TmpTxt.SelStart = 0
End If
TmpTxt.SelBold = False
TmpTxt.SelItalic = False
TmpTxt.SelUnderline = False
TmpTxt.SelFontName = "MS Sans Serif"
TmpTxt.SelText = CRLF
textmsg = ColourzScheme(2) & "[" & ColourzScheme(1) & Time & ColourzScheme(2) & "] " & ColourzScheme(10) & textmsg

Dim tmp1, tmp2, tmp3, tmp4
For i = 1 To Len(textmsg)
  If Mid(textmsg, i, 1) = "" Then
    TmpTxt.SelText = tmp3
    tmp3 = ""
    On Error GoTo ErrLoc
    tmp1 = Mid(textmsg, i + 1, 2)
RepeatLoc:
    If tmp1 > 15 Then
      tmp1 = tmp1 - 16
      GoTo RepeatLoc
    End If
    If tmp1 = 0 Then
      TmpTxt.SelColor = Colourz(16)
    Else
      TmpTxt.SelColor = Colourz(tmp1)
    End If
    i = i + 2
    GoTo EndLoc
ErrLoc:
    On Error GoTo EndLoc
    tmp1 = Mid(textmsg, i + 1, 1)
    If tmp1 = 0 Then
      TmpTxt.SelColor = Colourz(16)
    Else
      TmpTxt.SelColor = Colourz(tmp1)
    End If
    i = i + 1
EndLoc:
  ElseIf Mid(textmsg, i, 1) = Chr$(10) Or Mid(textmsg, i, 1) = Chr$(13) Or Mid(textmsg, i, 2) = Chr$(10) & Chr$(13) Or Mid(textmsg, i, 2) = Chr$(13) & Chr$(10) Then
  ElseIf Mid(textmsg, i, 1) = "" Then
    TmpTxt.SelText = tmp3
    tmp3 = ""
    TmpTxt.SelBold = Not TmpTxt.SelBold
  ElseIf Mid(textmsg, i, 4) = "rn" Or Mid(textmsg, i, 2) = "n" Then
    TmpTxt.SelText = tmp3
    tmp3 = ""
    Incoming.SelRTF = TmpTxt.TextRTF
    TmpTxt.Text = ""
    Incoming.SelText = CRLF
    tmp = Incoming.Font
    Incoming.SelBold = False
    Incoming.SelItalic = False
    Incoming.SelUnderline = False
    Incoming.SelFontName = "MS Sans Serif"
    Incoming.SelColor = RGB(0, 128, 255)
    Incoming.SelText = "Continued"
    Incoming.SelColor = RGB(255, 0, 0)
    Incoming.SelText = "> "
    Incoming.Font = tmp
    If Mid(textmsg, i, 4) = "rn" Then
      i = i + 3
    Else
      i = i + 1
    End If
  ElseIf Mid(textmsg, i, 1) = "" Then
    TmpTxt.SelText = tmp3
    tmp3 = ""
    TmpTxt.SelItalic = Not TmpTxt.SelItalic
  ElseIf Mid(textmsg, i, 1) = "" Then
    TmpTxt.SelText = tmp3
    tmp3 = ""
    TmpTxt.SelUnderline = Not TmpTxt.SelUnderline
  ElseIf Mid(textmsg, i, 1) = "" Then
    TmpTxt.SelText = tmp3
    tmp3 = ""
    If TmpTxt.SelFontName = "MS Sans Serif" Then
      TmpTxt.SelFontName = "Courier"
    Else
      TmpTxt.SelFontName = "MS Sans Serif"
    End If
  Else
    tmp3 = tmp3 & Mid(textmsg, i, 1)
  End If
  If i = Len(textmsg) Then
    TmpTxt.SelText = tmp3
  End If
Next i

If PauseOutput = False Then
  Incoming.SelStart = Len(Incoming.Text)
  Incoming.SelRTF = TmpTxt.TextRTF
  Incoming.SelStart = Len(Incoming.Text)
End If
End Sub

Private Sub Incoming_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PauseOutput = True
End Sub

Private Sub Incoming_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PauseOutput = False
If Len(TempTxt) > 0 Then
  Incoming.SelStart = Len(Incoming.Text)
  Incoming.SelRTF = TmpTxt.TextRTF
  TmpTxt.Text = ""
  Incoming.SelStart = Len(Incoming.Text)
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Tree.Height = ScaleHeight
Incoming.Width = ScaleWidth - Incoming.Left
Outgoing.Width = Incoming.Width
Incoming.Height = ScaleHeight - Outgoing.Height - 30
Outgoing.Top = Incoming.Height + 30
End Sub

Private Sub Outgoing_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then Outgoing.Text = LastMessage
End Sub

Private Sub Outgoing_KeyPress(KeyAscii As Integer)
Dim msg As String
Dim OrgMsg As String
If KeyAscii <> 13 Then Exit Sub
KeyAscii = 0
msg = Outgoing.Text
LastMessage = msg
OrgMsg = msg
If Left(UCase(msg), 3) = "/CN" Then
  SendData "NICK 'Vampirate\bBot"
  Outgoing.Text = ""
  Exit Sub
End If
If Left(UCase(msg), 5) = "/CHAR" Then
  SendData "DATA " & ChannelName & " CCUDI1 :# Appears as Hotaru"
  Outgoing.Text = ""
  Exit Sub
End If
If Left(UCase(msg), 7) = "/ACTION" Or Left(UCase(msg), 3) = "/ME" Then
  If InStr(msg, " ") = 0 Then
    Outgoing = ""
    Exit Sub
  End If
  msg = FadeColourz(Mid(msg, InStr(msg, " ") + 1))
  FrmDebug.AddText msg
  SendData "PRIVMSG " & ChannelName & " :ACTION " & msg & ""
  AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Action" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & Nickname & ColourzScheme(6) & "> " & ColourzScheme(9) & msg
  Outgoing.Text = ""
  Exit Sub
End If
If Left$(msg, 1) <> "/" Then
  If Len(msg) = 0 Then
    SendData "PRIVMSG " & ChannelName & " <Chr>"
    AddText "20<28" & Nickname & "20>16 <Chr>"
    Exit Sub
  End If
  msg2 = FadeColourz(msg)
  SendData "PRIVMSG " & ChannelName & " :06" & msg2
  AddText "20<28" & Nickname & "20>06 " & msg2
Else
  Outgoing.Text = Mid$(Outgoing.Text, 2)
  If InStr(Outgoing.Text, " ") = 0 Then Outgoing.Text = Outgoing.Text & " "
  msg = Mid$(Outgoing.Text, InStr(Outgoing.Text, " ") + 1)
  Select Case UCase$(Left$(Outgoing.Text, InStr(Outgoing.Text, " ") - 1))
    Case "ME"
      If NameList.ListCount > 0 Then SendData "PRIVMSG " & ChannelName & " :" & Chr$(1) & "ACTION " & msg & Chr$(1)
      AddText "* " & Nickname & " " & msg
    Case "MSG"
      SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " :" & Mid$(msg, InStr(msg, " ") + 1)
      AddText "=->" & Left$(msg, InStr(msg, " ") - 1) & "<-= " & Mid$(msg, InStr(msg, " ") + 1)
    Case "CTCP"
      On Error Resume Next
      If Trim(UCase(Right(OrgMsg, 7))) = "VERSION" Then
        SendData ("PRIVMSG " & Left(msg, InStr(msg, " ")) & "VERSION")
        AddText "*** Sending CTCP Version To: " & Left(msg, InStr(msg, " "))
      End If
    Case Else
      If InStr(OrgMsg, " ") = 0 Then OrgMsg = OrgMsg & " "
      SendData (UCase(Mid(OrgMsg, 2, InStr(OrgMsg, " ") - 2))) & " " & Right(OrgMsg, Len(OrgMsg) - InStr(OrgMsg, " "))
  End Select
End If
Outgoing.Text = ""
End Sub
