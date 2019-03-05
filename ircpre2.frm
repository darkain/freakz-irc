VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmRoom 
   Caption         =   "DooVoo Chat - Alpha"
   ClientHeight    =   4170
   ClientLeft      =   1545
   ClientTop       =   1380
   ClientWidth     =   6735
   Icon            =   "ircpre2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   6735
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox Topic 
      Height          =   305
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   529
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"ircpre2.frx":030A
   End
   Begin VB.PictureBox Label1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5160
      MousePointer    =   9  'Size W E
      ScaleHeight     =   735
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   240
      Width           =   45
   End
   Begin RichTextLib.RichTextBox TmpTxt 
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"ircpre2.frx":038C
   End
   Begin RichTextLib.RichTextBox Incoming 
      Height          =   3495
      Left            =   0
      TabIndex        =   2
      Top             =   315
      Width           =   5232
      _ExtentX        =   9234
      _ExtentY        =   6165
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"ircpre2.frx":040E
   End
   Begin VB.ListBox NameList 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   3840
      IntegralHeight  =   0   'False
      ItemData        =   "ircpre2.frx":0490
      Left            =   5250
      List            =   "ircpre2.frx":0492
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   315
      Width           =   1488
   End
   Begin VB.TextBox Outgoing 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Width           =   5232
   End
End
Attribute VB_Name = "FrmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ChannelName As String
Dim LastMessage
Dim PauseOutput As Byte

Function GetChannelName() As String
GetChannelName = ChannelName
End Function

Sub SetChannelName(NewName As String)
ChannelName = NewName
Caption = ChannelName
End Sub

Sub AddText(textmsg As String)
  Dim i As Integer
  Dim tmp1 As String, tmp2 As String
  Dim tmp3 As String, tmp As String

  If PauseOutput = False Then
    TmpTxt.Text = ""
    TmpTxt.SelStart = 0
  End If
  
  TmpTxt.SelBold = False
  TmpTxt.SelItalic = False
  TmpTxt.SelUnderline = False
  TmpTxt.SelFontName = "MS Sans Serif"
  TmpTxt.SelText = "" 'CRLF
  textmsg = ColourzScheme(2) & "[" & ColourzScheme(1) & Time & ColourzScheme(2) & "] " & ColourzScheme(10) & textmsg

  For i = 1 To Len(textmsg)
    If Mid(textmsg, i, 1) = "" Then
      TmpTxt.SelText = tmp3
      tmp3 = ""
      tmp1 = Val(Mid(textmsg, i + 1, 2))
      tmp2 = tmp1
RepeatLoc:
      On Error GoTo ErrLoc
      
      If tmp1 > 15 Then
        tmp1 = tmp1 - 16
        GoTo RepeatLoc
      End If
      
      If tmp1 = 0 Then
        TmpTxt.SelColor = Colourz(16)
      Else
        TmpTxt.SelColor = Colourz(tmp1)
      End If
      
      If Mid(textmsg, i + 1, 1) = "0" Or tmp2 > 9 Then
        i = i + 2
      Else
        i = i + 1
      End If
      GoTo EndLoc
ErrLoc:
      On Error GoTo EndLoc
      tmp1 = Val(Mid(textmsg, i + 1, 1))
      If tmp1 = "0" Then
        TmpTxt.SelColor = Colourz(16)
      Else
        TmpTxt.SelColor = Colourz(tmp1)
      End If
      i = i + 1
EndLoc:
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
    
    TmpTxt.Text = ""
  End If
End Sub

Sub SetTopic(textmsg As String)
Dim i As Integer

Topic.Text = ""
Topic.SelBold = False
Topic.SelItalic = False
Topic.SelUnderline = False
Topic.SelFontName = "MS Sans Serif"
textmsg = ColourzScheme(9) & textmsg

Dim tmp1, tmp2, tmp3, tmp4
For i = 1 To Len(textmsg)
  If Mid(textmsg, i, 1) = "" Then
    Topic.SelText = tmp3
    tmp3 = ""
    On Error GoTo ErrLoc
    tmp1 = Mid(textmsg, i + 1, 2)
RepeatLoc:
    If tmp1 > 15 Then
      tmp1 = tmp1 - 16
      GoTo RepeatLoc
    End If
    If tmp1 = 0 Then
      Topic.SelColor = Colourz(16)
    Else
      Topic.SelColor = Colourz(tmp1)
    End If
    i = i + 2
    GoTo EndLoc
ErrLoc:
    On Error GoTo EndLoc
    tmp1 = Mid(textmsg, i + 1, 1)
    If tmp1 = 0 Then
      Topic.SelColor = Colourz(16)
    Else
      Topic.SelColor = Colourz(tmp1)
    End If
    i = i + 1
EndLoc:
  ElseIf Mid(textmsg, i, 1) = Chr$(10) Or Mid(textmsg, i, 1) = Chr$(13) Then
  ElseIf Mid(textmsg, i, 1) = "" Then
    Topic.SelText = tmp3
    tmp3 = ""
    Topic.SelBold = Not TmpTxt.SelBold
  ElseIf Mid(textmsg, i, 1) = "" Then
    Topic.SelText = tmp3
    tmp3 = ""
    Topic.SelBold = Not TmpTxt.SelBold
  ElseIf Mid(textmsg, i, 1) = "" Then
    Topic.SelText = tmp3
    tmp3 = ""
    Topic.SelItalic = Not TmpTxt.SelItalic
  ElseIf Mid(textmsg, i, 1) = "" Then
    Topic.SelText = tmp3
    tmp3 = ""
    Topic.SelUnderline = Not TmpTxt.SelUnderline
  ElseIf Mid(textmsg, i, 1) = "" Then
    Topic.SelText = tmp3
    tmp3 = ""
    If Topic.SelFontName = "MS Sans Serif" Then
      Topic.SelFontName = "Courier"
    Else
      Topic.SelFontName = "MS Sans Serif"
    End If
  Else
    tmp3 = tmp3 & Mid(textmsg, i, 1)
  End If
  If i = Len(textmsg) Then
    Topic.SelText = tmp3
  End If
Next i
End Sub

Sub SendData(textmsg As String)
FrmMain.SendData textmsg
End Sub

Private Sub Form_Load()
Incoming.SelColor = NameList.ForeColor
Me.Width = FrmMain.Width / 1.25
Me.Height = FrmMain.Height / 1.25
Me.Left = FrmMain.Width / 2 - Me.Width / 2
Me.Top = FrmMain.Height / 2 - Me.Height / 2
Label1.Top = Incoming.Top
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FullUnload = True Then Exit Sub
SendData "PART " & ChannelName
ChannelName = ""
NameList.Clear
Incoming.Text = ""
Me.Hide
Cancel = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
Topic.Width = Me.Width - 140
Incoming.Width = Me.Width - NameList.Width - 165
NameList.Left = Me.Width - NameList.Width - 125
Outgoing.Width = Incoming.Width
Incoming.Height = Me.Height - Topic.Height - Outgoing.Height - 445
Outgoing.Top = Me.Height - Outgoing.Height - 405
NameList.Height = Me.Height - 715
Label1.Height = Incoming.Height
Label1.Left = NameList.Left - Label1.Width - 30
End Sub

Private Sub Incoming_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PauseOutput = True
End Sub

Private Sub Incoming_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PauseOutput = False
If Len(TmpTxt.Text) > 0 Then
  Incoming.SelStart = Len(Incoming.Text)
  Incoming.SelRTF = TmpTxt.TextRTF
  TmpTxt.Text = ""
  Incoming.SelStart = Len(Incoming.Text)
End If
End Sub

Private Sub NameList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 2 Then Exit Sub
  Dim i As Integer
  For i = 1 To MaxChanNum
    If Channels(i).GetChannelName = Me.GetChannelName Then
      ChannelMnu = i
      Exit For
    End If
  Next i
  PopupMenu FrmMain.MnuNickList
End Sub

Private Sub Outgoing_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then Outgoing.Text = LastMessage
End Sub

Private Sub Outgoing_KeyPress(KeyAscii As Integer)
Dim msg As String, msg2 As String
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
  'msg2 = msg
  SendData "PRIVMSG " & ChannelName & " :" & "3 " & msg2
  AddText "20<28" & Nickname & "20> " & "3 " & msg2
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

Private Sub Topic_GotFocus()
Outgoing.SetFocus
End Sub
