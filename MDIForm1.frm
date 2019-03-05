VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H00000000&
   Caption         =   "DooVoo Chat - Alpha"
   ClientHeight    =   4845
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   7245
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer CTCP_Timer 
      Interval        =   2000
      Left            =   0
      Top             =   3960
   End
   Begin MSWinsockLib.Winsock TCP1 
      Left            =   0
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu FileConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu FileSetup 
      Caption         =   "&Setup"
   End
   Begin VB.Menu MnuWindowList 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu MnuWinCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu MnuWinTileH 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu MnuWinTileV 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu MnuArangeIcon 
         Caption         =   "Arange Icons"
      End
   End
   Begin VB.Menu HelpAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu MnuNickList 
      Caption         =   "NickListMnu"
      Visible         =   0   'False
      Begin VB.Menu MnuNickWhois 
         Caption         =   "Whois"
      End
      Begin VB.Menu MnuDNS 
         Caption         =   "DNS"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMode 
         Caption         =   "Mode"
         Begin VB.Menu MnuOp 
            Caption         =   "Op"
         End
         Begin VB.Menu MnuDeOp 
            Caption         =   "DeOp"
         End
         Begin VB.Menu MnuSep2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuOwner 
            Caption         =   "Owner"
         End
         Begin VB.Menu MnuDeowner 
            Caption         =   "DeOwner"
         End
         Begin VB.Menu MnuSep3 
            Caption         =   "-"
         End
         Begin VB.Menu MnuVoice 
            Caption         =   "Voice"
         End
         Begin VB.Menu MnuDevoice 
            Caption         =   "DeVoice"
         End
      End
      Begin VB.Menu MnuKick 
         Caption         =   "Kick"
         Begin VB.Menu MnuKickKick 
            Caption         =   "Kick"
         End
         Begin VB.Menu MnuKickReason 
            Caption         =   "Kick Reason"
         End
         Begin VB.Menu MnuKickBan 
            Caption         =   "Kick-Ban"
         End
         Begin VB.Menu MnuKickManReason 
            Caption         =   "Kick-Ban Reason"
         End
         Begin VB.Menu MnuBan 
            Caption         =   "Ban"
         End
      End
      Begin VB.Menu MnuAddLists 
         Caption         =   "Add To List"
         Begin VB.Menu MnuAddShitList 
            Caption         =   "Shit List"
         End
         Begin VB.Menu MnuAddSlienceList 
            Caption         =   "Silence List"
         End
         Begin VB.Menu MnuAddDeopList 
            Caption         =   "DeOp List"
         End
         Begin VB.Menu MnuAddOpList 
            Caption         =   "Op List"
         End
         Begin VB.Menu MnuAddOwnerList 
            Caption         =   "Owner List"
         End
         Begin VB.Menu MnuAddProtect 
            Caption         =   "Protect List"
         End
      End
      Begin VB.Menu MnuRemoveList 
         Caption         =   "Remove From List"
         Begin VB.Menu MnuRemoveShitlist 
            Caption         =   "Shit List"
         End
         Begin VB.Menu MnuRemoveSlienceList 
            Caption         =   "Slience"
         End
         Begin VB.Menu MnuRemoveDeopList 
            Caption         =   "DeOp List"
         End
         Begin VB.Menu MnuRemoveOpList 
            Caption         =   "Op List"
         End
         Begin VB.Menu MnuRemoveOwnerList 
            Caption         =   "Owner List"
         End
         Begin VB.Menu MnuRemoveProtectList 
            Caption         =   "Protect List"
         End
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CTCP_Reply As Byte

Sub OnConnect()
SendData "IRCX"
SendData "JOIN #LP"
End Sub

Sub SendData(textmsg As String)
On Error Resume Next
TCP1.SendData textmsg & CRLF
End Sub

Sub AddText(textmsg As String, WindowName As String)
For i = 0 To MaxChanNum
  If UCase(WindowName) = "STATUS" Then
    StatusWin.AddText textmsg
    Exit For
  ElseIf UCase(WindowName) = "DEBUG" Then
    FrmDebug.AddText textmsg
    Exit For
  ElseIf UCase(WindowName) = UCase(Nickname) Then
    FrmPrivMsg.AddText textmsg
    Exit For
  ElseIf UCase(WindowName) = UCase(Channels(i).GetChannelName) Then
    Channels(i).AddText textmsg
    Exit For
  End If
Next i
End Sub

Private Sub CTCP_Timer_Timer()
CTCP_Reply = True
End Sub

Private Sub FileConnect_Click()
If FileConnect.Caption = "&Connect" Then
  On Error GoTo ErrHan
  TCP1.Close
  TCP1.LocalPort = PortNum
  TCP1.RemoteHost = ServerAddress
  TCP1.RemotePort = Port
  TCP1.Connect
  AddText "20[28Connection20] 32Connecting to " & ServerAddress, "STATUS"
  FileConnect.Caption = "&Disconnect"
Else
  TCP1_Close
End If
Exit Sub

ErrHan:
  If PortNum = 17000 Then Exit Sub
  TCP1.Close
  PortNum = PortNum + 1
  TCP1.LocalPort = PortNum
Resume
End Sub

Private Sub HelpAbout_Click()
Load FrmAbout
FrmAbout.Visible = True
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FullUnload = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
TCP1.Close
End Sub

Private Sub MnuArangeIcon_Click()
FrmMain.Arrange 3
End Sub

Private Sub MnuDeOp_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    SendData "MODE " & Channels(ChannelMnu).GetChannelName & " -o " & tmp1
  End If
Next i
End Sub

Private Sub MnuDeowner_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    SendData "MODE " & Channels(ChannelMnu).GetChannelName & " -q " & tmp1
  End If
Next i
End Sub

Private Sub MnuDevoice_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    SendData "MODE " & Channels(ChannelMnu).GetChannelName & " -v " & tmp1
  End If
Next i
End Sub

Private Sub MnuDNS_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    AddText tmp1, "STATUS"
  End If
Next i
End Sub

Private Sub FileSetup_Click()
Load FrmSetup
FrmSetup.Visible = True
End Sub

Private Sub MnuNickWhois_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    SendData "WHOIS " & tmp1
  End If
Next i
End Sub

Private Sub MnuOp_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    SendData "MODE " & Channels(ChannelMnu).GetChannelName & " +o " & tmp1
  End If
Next i
End Sub

Private Sub MnuOwner_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    SendData "MODE " & Channels(ChannelMnu).GetChannelName & " +q " & tmp1
  End If
Next i
End Sub

Private Sub MDIForm_Load()
Dim erg, X, Y
On Error Resume Next
SetColourz
Me.Width = Screen.Width / 1.25
Me.Height = Screen.Height / 1.25
Me.Left = Screen.Width / 2 - Me.Width / 2
'Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Top = Screen.Height + 1000
CRLF = Chr$(13) & Chr$(10)
PortNum = 6667
ServerAddress = "192.168.0.1"
Port = 6667
Nickname = "Darkain_X"
Load StatusWin
StatusWin.Visible = True

Load FrmPics
Y = 0
For i1 = 0 To Int(Screen.Height / FrmPics.Picture1.Height) + 1
  X = 0
  For i2 = 0 To Int(Screen.Width / FrmPics.Picture1.Width) + 1
    erg = BitBlt(FrmPics.Picture2.hDC, X, Y, FrmPics.Picture1.Width - 1, FrmPics.Picture1.ScaleHeight - 1, FrmPics.Picture1.hDC, 0, 0, SRCCOPY)
    X = X + FrmPics.Picture1.ScaleWidth - 1
  Next i2
  Y = Y + FrmPics.Picture1.ScaleHeight - 1
Next i1
Me.Picture = FrmPics.Picture2.Image
For i = 0 To MaxChanNum
  Channels(i).Hide
Next i
Unload FrmPics
Load FrmDebug
FrmDebug.Width = 5000
FrmDebug.Height = 5000
FrmDebug.Left = Me.ScaleWidth - FrmDebug.Width
FrmDebug.Top = Me.ScaleHeight - FrmDebug.Height
FrmDebug.Show
Load FrmPrivMsg
FrmPrivMsg.Width = 5000
FrmPrivMsg.Height = 5000
FrmPrivMsg.Left = Me.ScaleWidth - FrmDebug.Width - 4985
FrmPrivMsg.Top = Me.ScaleHeight - FrmDebug.Height
FrmPrivMsg.Show
StatusWin.Outgoing.SetFocus
Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub MnuVoice_Click()
Dim tmp1 As String
For i = 0 To Channels(ChannelMnu).NameList.ListCount - 1
  If Channels(ChannelMnu).NameList.Selected(i) = True Then
    tmp1 = Channels(ChannelMnu).NameList.List(i)
    If Left(tmp1, 1) = "." Or Left(tmp1, 1) = "@" Or Left(tmp1, 1) = "+" Then
      tmp1 = Mid(tmp1, 2)
    End If
    If Left(tmp1, 1) = "'" Then
      tmp1 = Mid(tmp1, 2)
    End If
    SendData "MODE " & Channels(ChannelMnu).GetChannelName & " +v " & tmp1
  End If
Next i
End Sub

Private Sub MnuWinCascade_Click()
FrmMain.Arrange 0
End Sub

Private Sub MnuWinTileH_Click()
FrmMain.Arrange 1
End Sub

Private Sub MnuWinTileV_Click()
FrmMain.Arrange 2
End Sub

Private Sub TCP1_Close()
FrmMain.FileConnect.Caption = "&Connect"
AddText "20[28Connection20] 32Disconnected", "STATUS"
TCP1.Close
ConnectedToServer = False
For i = 0 To MaxChanNum
  Channels(i).Incoming.Text = ""
  Channels(i).NameList.Clear
  Channels(i).SetChannelName ""
  Channels(i).Hide
Next i
End Sub

Private Sub TCP1_Connect()
AddText "20[28Connection20] 32Connected", "STATUS"
AddText "20[28Connection20] 32Sending Login Information", "STATUS"
SendData "NICK " & Nickname
SendData "USER FreakZ " & TCP1.LocalIP & " " & ServerAddress & " :FreakZ"
End Sub

Private Sub TCP1_DataArrival(ByVal bytesTotal As Long)
Dim inData As String
Dim sline As String
Dim msg As String
Dim msg2 As String
Dim CommandStart As Integer
Dim CommandLength As Integer
Dim CommandName As String
Dim CommandNick As String
Dim CommandChannel As String
Dim ChannelNum As Integer
Dim X
Dim tmp1, tmp2, tmp3, tmp4
TCP1.GetData inData, vbString
inData = OldText & inData
X = 0
If Right$(inData, 2) = CRLF Then X = 1
If Right$(inData, 1) = Chr$(10) Then X = 1
If Right$(inData, 1) = Chr$(13) Then X = 1
If X = 1 Then
  OldText = ""
Else
  OldText = inData: Exit Sub
End If
  
again:
  GoSub parsemsg ' get next msg fragment
  If Left$(sline, 6) = "PING :" Then
    AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Ping" & ColourzScheme(4) & "] " & ColourzScheme(9) & "Pong", "STATUS"
    SendData "PONG " & ServerName
    GoTo again
  End If
  If Left$(sline, 5) = "ERROR" Then
    tmp1 = Mid$(sline, InStr(sline, "("))
    tmp2 = Left(tmp1, Len(tmp1) - 1)
    AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Error" & ColourzScheme(4) & "] " & ColourzScheme(9) & tmp2, "STATUS"
    GoTo again
  End If
  On Error Resume Next
  CommandStart = InStr(1, sline, " ")
  CommandLength = InStr(CommandStart + 1, sline, " ") - CommandStart
  CommandName = Mid(sline, CommandStart + 1, CommandLength - 1)
  CommandNick = Mid$(sline, 2, InStr(sline, "!") - 2)
  CommandIP = Mid(sline, InStr(sline, "@"), InStr(sline, " ") - InStr(sline, "@"))
  tmp1 = InStr(sline, "#")
  tmp2 = 1
  For i = tmp1 To 1 Step -1
    If Mid(sline, i, 1) = " " Then
      tmp2 = i + 1
      Exit For
    End If
  Next i
  If Mid(sline, tmp2, 1) = ":" Then tmp2 = tmp2 + 1
  tmp3 = InStr(tmp2 + 1, sline, " ")
  If tmp3 = 0 Then
    tmp4 = Len(sline)
  Else
    tmp4 = tmp3 - tmp2
  End If
  CommandChannel = Mid(sline, tmp2, tmp4)
  ChannelNum = -1
  For i = 0 To MaxChanNum
    If UCase(Channels(i).GetChannelName) = UCase(CommandChannel) Then
      ChannelNum = i
      Exit For
    End If
  Next i
  AddText "SLINE: " & sline, "DEBUG"
  

  Select Case CommandName
    Case "001"   ' Connected
      ServerName = Mid$(sline, 2, InStr(sline, " ") - 2)
      'OnConnect
      GoTo again
    
'    Case 250 To 270 ' Ignore
'      GoTo again
    
    Case 372 To 375  ' MOTD
      AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "MOTD" & ColourzScheme(4) & "] " & ColourzScheme(9) & Mid(sline, InStr(2, sline, ":") + 1), "STATUS"
      GoTo again
    
    Case 376 ' End MOTD
      AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "MOTD" & ColourzScheme(4) & "] " & ColourzScheme(9) & Mid(sline, InStr(2, sline, ":") + 1), "STATUS"
      If ConnectedToServer = False Then
        AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "TCP" & ColourzScheme(4) & "] " & ColourzScheme(8) & "<" & ColourzScheme(7) & "Local IP" & ColourzScheme(4) & "> " & ColourzScheme(9) & TCP1.LocalIP, "STATUS"
        AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "TCP" & ColourzScheme(4) & "] " & ColourzScheme(8) & "<" & ColourzScheme(7) & "Local Port" & ColourzScheme(4) & "> " & ColourzScheme(9) & TCP1.LocalPort, "STATUS"
        AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "TCP" & ColourzScheme(4) & "] " & ColourzScheme(8) & "<" & ColourzScheme(7) & "Remote IP" & ColourzScheme(4) & "> " & ColourzScheme(9) & TCP1.RemoteHostIP, "STATUS"
        AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "TCP" & ColourzScheme(4) & "] " & ColourzScheme(8) & "<" & ColourzScheme(7) & "Remote Address" & ColourzScheme(4) & "> " & ColourzScheme(9) & TCP1.RemoteHost, "STATUS"
        AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "TCP" & ColourzScheme(4) & "] " & ColourzScheme(8) & "<" & ColourzScheme(7) & "Remote Port" & ColourzScheme(4) & "> " & ColourzScheme(9) & TCP1.RemotePort, "STATUS"
      End If
      OnConnect
      ConnectedToServer = True
      GoTo again
    
    Case 332 ' Topic
      For i = 0 To MaxChanNum
        If UCase(CommandChannel) = UCase(Channels(i).GetChannelName) Then
          Channels(i).SetTopic Mid$(sline, InStr(2, sline, ":") + 1)
          Exit For
        End If
      Next i
      GoTo again
    
    Case 353 ' Name list
      msg = Mid$(sline, InStr(2, sline, ":") + 1)
      For i = 0 To MaxChanNum
        If Channels(i).GetChannelName = "" Then
          Channels(i).SetChannelName CommandChannel
          Channels(i).Show
          ChannelNum = i
          Exit For
        End If
        If i = MaxChanNum And ChannelNum <> MaxChanNum Then
          AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Server" & ColourzScheme(4) & "] " & ColourzScheme(9) & "Too many channels open at once", "STATUS"
          SendData "PART " & CommandChannel
        End If
      Next i
      Do Until msg = ""
        X = InStr(msg, " ")
        If X <> 0 Then
          Channels(ChannelNum).NameList.AddItem Left$(msg, X - 1)
          msg = Mid$(msg, X + 1)
        Else
          Channels(ChannelNum).NameList.AddItem msg
          msg = ""
        End If
      Loop
      GoTo again
    
    Case 366 ' End of Name List
      SendData "DATA " & CommandChannel & " CCUDI1 :# Appears as Hotaru"
      GoTo again
    
    Case 800 ' IRCX Mode
      AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "IRCX" & ColourzScheme(4) & "] " & ColourzScheme(9) & Mid(sline, 2), "STATUS"
      GoTo again
  
    Case "JOIN" ' User Joins A Channel
      If CommandNick <> Nickname Then
        Channels(ChannelNum).NameList.AddItem CommandNick
        AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Join" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & CommandIP & ColourzScheme(6) & ">", CommandChannel
      End If
      GoTo again
  
    Case "PART" ' User Leaves A Channel
      If ChannelNum <> -1 Then
        For i = 0 To Channels(ChannelNum).NameList.ListCount
          If Right$(Channels(ChannelNum).NameList.List(i), Len(CommandNick)) = CommandNick Then
            Channels(i1).NameList.RemoveItem i
            AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Part" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & CommandIP & ColourzScheme(6) & ">", CommandChannel
            Exit For
          End If
        Next i
      End If
      GoTo again
  
    Case "KICK" ' User Kicked From Channel
      AddText "20[28Kick20] <28" & CommandNick & "20> 32kicked " & Mid(sline, InStr(sline, CommandChannel) + Len(CommandChannel) + 1, (InStr(2, sline, ":") - InStr(sline, CommandChannel)) - Len(CommandChannel) - 2) & " : " & Mid(sline, InStr(2, sline, ":") + 1), CommandChannel
      For i = 0 To Channels(ChannelNum).NameList.ListCount
        If Right$(Channels(ChannelNum).NameList.List(i), Len(CommandNick)) = CommandNick Then
          Channels(ChannelNum).NameList.RemoveItem i
          Exit For
        End If
      Next i
      GoTo again

    Case "QUIT" ' User Disconnects From Server
      For i1 = 0 To MaxChanNum
        If Channels(i1).GetChannelName <> "" Then
          For i2 = 0 To Channels(i1).NameList.ListCount
            If Right$(Channels(i1).NameList.List(i2), Len(CommandNick)) = CommandNick Then
              Channels(i1).NameList.RemoveItem i2
              AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Quit" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & CommandIP & ColourzScheme(6) & "> " & ColourzScheme(9) & Mid(sline, InStr(2, sline, ":") + 1), Channels(i1).GetChannelName
            End If
          Next i2
        End If
      Next i1
      GoTo again

    Case "KILL" ' User K-Lined From Server
      For i1 = 0 To MaxChanNum
        If Channels(i1).GetChannelName <> "" Then
          For i2 = 0 To Channels(i1).NameList.ListCount
            If Right$(Channels(i1).NameList.List(i2), Len(CommandNick)) = CommandNick Then
              Channels(i1).NameList.RemoveItem i2
              AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Kill" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & Mid(sline, InStr(2, sline, ":") + 1), Channels(i1).GetChannelName
            End If
          Next i2
        End If
      Next i1
      GoTo again

    Case "MODE" ' Mode Change
      tmp1 = InStr(CommandStart + 6, sline, " ")
      tmp2 = InStr(tmp1 + 1, sline, " ") - tmp1
      tmp3 = InStr(tmp1 + tmp2, sline, " ")
      tmp4 = Mid(sline, tmp3 + 1)
      tmp5 = Mid(sline, tmp1 + 1, tmp2 - 1)
      AddText "28" & CommandNick & " 20Set Mode: 25" & tmp5 & " 28" & tmp4, CommandChannel
      For i = 0 To Channels(ChannelNum).NameList.ListCount
        If Left(Channels(ChannelNum).NameList.List(i), 1) = "." Or Left(Channels(ChannelNum).NameList.List(i), 1) = "@" Or Left(Channels(ChannelNum).NameList.List(i), 1) = "+" Then
          If Right$(Channels(ChannelNum).NameList.List(i), Len(tmp4)) = tmp4 Then
            Select Case tmp5
              Case "+q"
                Channels(ChannelNum).NameList.RemoveItem i
                Channels(ChannelNum).NameList.AddItem "." & tmp4
                If tmp4 = Nickname Then
                  'SendData "ACCESS " & CommandChannel & " ADD OWNER @" & TCP1.LocalIP
                End If
              Case "+o" Or "+a"
                Channels(ChannelNum).NameList.RemoveItem i
                Channels(ChannelNum).NameList.AddItem "@" & tmp4
              Case "+v"
                If Left(Channels(ChannelNum).NameList.List(i), 1) = "." Or Left(Channels(ChannelNum).NameList.List(i), 1) = "@" Then Exit For
                Channels(ChannelNum).NameList.RemoveItem i
                Channels(ChannelNum).NameList.AddItem "+" & tmp4
              Case "-v"
                If Left(Channels(ChannelNum).NameList.List(i), 1) = "." Or Left(Channels(ChannelNum).NameList.List(i), 1) = "@" Then Exit For
                Channels(ChannelNum).NameList.RemoveItem i
                Channels(ChannelNum).NameList.AddItem tmp4
              Case Else
                Channels(ChannelNum).NameList.RemoveItem i
                Channels(ChannelNum).NameList.AddItem tmp4
            End Select
            Exit For
          End If
        Else
          If Channels(ChannelNum).NameList.List(i) = tmp4 Then
            Channels(ChannelNum).NameList.RemoveItem i
            Select Case tmp5
              Case "+q"
                Channels(ChannelNum).NameList.AddItem "." & tmp4
                If tmp4 = Nickname Then
                  'SendData "ACCESS " & CommandChannel & " ADD OWNER @" & TCP1.LocalIP
                End If
              Case "+o"
                Channels(ChannelNum).NameList.AddItem "@" & tmp4
              'Case "-o"
              '  Channels(ChannelNum).NameList.AddItem "+" & tmp4
              Case "+v"
                If Left(Channels(ChannelNum).NameList.List(i), 1) = "." Or Left(Channels(ChannelNum).NameList.List(i), 1) = "@" Then Exit For
                Channels(ChannelNum).NameList.AddItem "+" & tmp4
              Case "-v"
                If Left(Channels(ChannelNum).NameList.List(i), 1) = "." Or Left(Channels(ChannelNum).NameList.List(i), 1) = "@" Then Exit For
                Channels(ChannelNum).NameList.AddItem tmp4
              Case Else
                Channels(ChannelNum).NameList.AddItem tmp4
            End Select
            Exit For
          End If
        End If
      Next i
      GoTo again
  
    Case "TOPIC" ' Channel Topic
      For i = 0 To MaxChanNum
        If UCase(CommandChannel) = UCase(Channels(i).GetChannelName) Then
          Channels(i).SetTopic Mid$(sline, InStr(2, sline, ":") + 1)
          Exit For
        End If
      Next i
      GoTo again

    Case "DATA" ' Misc Data
      GoTo again
  
    Case "NOTICE" ' Used Sends A Notice
      AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Notice" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & Mid(sline, InStr(2, sline, ":") + 1), "STATUS"
      GoTo again

    Case "NICK" ' User Changes Nickname
      If Nickname = CommandNick Then Nickname = Mid(sline, InStr(2, sline, ":") + 1)
      For i1 = 0 To MaxChanNum
        If Channels(i1).GetChannelName <> "" Then
          For i2 = 0 To Channels(i1).NameList.ListCount
            If Right$(Channels(i1).NameList.List(i2), Len(CommandNick)) = CommandNick Then
              tmp1 = Left(Channels(i1).NameList.List(i2), Len(Channels(i1).NameList.List(i2)) - Len(CommandNick))
              Channels(i1).NameList.RemoveItem i2
              Channels(i1).NameList.AddItem tmp1 & Mid(sline, InStr(2, sline, ":") + 1)
              AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Nick" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & Mid(sline, InStr(2, sline, ":") + 1), Channels(i1).GetChannelName
            End If
          Next i2
        End If
      Next i1
      GoTo again
    
    Case "PRIVMSG" ' Misc ProvMsg Commands
      tmp1 = Mid(sline, InStr(2, sline, ":") + 1)
      If Left(tmp1, 1) = "" And Right(tmp1, 1) = "" Then
        If UCase(tmp1) = "VERSION" Then
          If CTCP_Reply = True Then
            'SendData "NOTICE " & CommandNick & " :VERSION FreakZ IRC Client - Alpha"
            SendData "NOTICE " & CommandNick & " :VERSION Microsoft Chat 3.0 (4.71.3521) (text mode)"
            CTCP_Reply = False
            AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "CTCP" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & "Version " & ColourzScheme(8) & "(" & ColourzScheme(7) & "Replied" & ColourzScheme(8) & ")", "STATUS"
          Else
            AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "CTCP" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & "Version " & ColourzScheme(8) & "(" & ColourzScheme(7) & "Not Replied" & ColourzScheme(8) & ")", "STATUS"
          End If
        ElseIf UCase(Left(tmp1, 6)) = "AWAY" Then
          AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Return" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & ">", CommandChannel
        ElseIf UCase(Left(tmp1, 5)) = "AWAY" And UCase(Right(tmp1, 1)) = "" Then
          AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Away" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & Mid(tmp1, 7, Len(tmp1) - 7), CommandChannel
        ElseIf UCase(Left(tmp1, 6)) = "SOUND" And UCase(Right(tmp1, 1)) = "" Then
          AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Sound" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & Mid(tmp1, InStr(8, tmp1, " ") + 1, Len(Mid(tmp1, InStr(8, tmp1, " "))) - 2) & " 20(28" & Mid(tmp1, 8, InStr(8, tmp1, " ") - 8) & "20)", CommandChannel
        ElseIf UCase(Left(tmp1, 17)) = "ACTION THINKS:  " And UCase(Right(tmp1, 1)) = "" Then
          AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Thought" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(9) & "Thinks " & Mid(tmp1, 18, Len(tmp1) - 18), CommandChannel
        ElseIf UCase(Left(tmp1, 7)) = "ACTION" And UCase(Right(tmp1, 1)) = "" Then
          AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "Action" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & ">" & ColourzScheme(9) & Mid(tmp1, 9, Len(tmp1) - 9), CommandChannel
        Else
          If InStr(tmp1, " ") = 0 Then
            AddText ColourzScheme(4) & "[" & ColourzScheme(3) & "CTCP" & ColourzScheme(4) & "] " & ColourzScheme(6) & "<" & ColourzScheme(5) & CommandNick & ColourzScheme(6) & "> " & ColourzScheme(8) & "(28" & Mid(tmp1, 2, Len(tmp1) - 2) & "20)", "STATUS"
          Else
            AddText "20[28CTCP20] <28" & CommandNick & "20> 32" & Mid(tmp1, InStr(1, tmp1, " ") + 1, Len(Mid(tmp1, InStr(1, tmp1, " "))) - 2) & " 20(28" & Mid(tmp1, 2, InStr(1, tmp1, " ") - 2) & "20)", "STATUS"
          End If
        End If
      Else
        tmp1 = InStr(sline, "PRIVMSG") + 1
        tmp2 = InStr(tmp1, sline, " ") + 1
        tmp3 = InStr(tmp2 + 1, sline, " ")
        AddText "20<28" & CommandNick & CommandIP & "20> 00" & Mid$(sline, InStr(2, sline, ":") + 1), Mid$(sline, tmp2, tmp3 - tmp2)
      End If
      GoTo again
  End Select
  
  If Mid(sline, 2, Len(ServerName)) = ServerName Then
    AddText "20[28Server20] 32" & Mid(sline, InStr(2, sline, ":") + 1), "STATUS"
    GoTo again
  End If
  Resume
  GoTo again
Exit Sub


parsemsg:
  If inData = "" Then Exit Sub
  X = InStr(inData, CRLF)
  If X <> 0 Then
    sline = Left$(inData, X - 1)
    If Len(inData) > X + 2 Then
      inData = Mid$(inData, X + 2)
    Else
      inData = ""
    End If
  Else
    X = InStr(inData, Chr$(13))
    If X = 0 Then
      X = InStr(inData, Chr$(10))
    End If
    If X <> 0 Then
      sline = Left$(inData, X - 1)
    Else
      sline = inData
    End If
    If Len(inData) > X + 1 Then
      inData = Mid$(inData, X + 1)
    Else
      inData = ""
    End If
  End If
Return
End Sub

Private Sub TCP1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
AddText "Number: " & Number & "   Description: " & Description & "   Scode: " & Scode, "DEBUG"
If PortNum = 17000 Then
  AddText "Error: " & Number & " " & Description, "STATUS"
  Exit Sub
End If
PortNum = PortNum + 1
TCP1.Close
TCP1.LocalPort = PortNum
TCP1.Connect
End Sub
