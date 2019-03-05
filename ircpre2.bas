Attribute VB_Name = "Module1"
Option Explicit

DefInt A-Z
Global Const MaxChanNum As Integer = 1
Global FullUnload As Byte

'Global variables
Global ServerAddress As String ' Host Server Address
Global ServerName As String    ' Host Server Name
Global Port As String          ' Host Server Port
Global Nickname As String      ' My Nickname
Global Channels(MaxChanNum) As New FrmRoom
Global ChannelMnu As Integer
Global PortNum As Integer

Global CRLF As String
Global OldText As String
Global ConnectedToServer As Byte

Public Const SRCAND = &H8800C6    ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020   ' (DWORD) dest = source
Public Const SRCERASE = &H440328  ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046 ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086  ' (DWORD) dest = source OR dest
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global Colourz(16) As ColorConstants
Global ColourzScheme(16) As String
Global ColourStyle As Byte
      
Sub SetColourz()
Colourz(1) = RGB(96, 96, 96)
Colourz(2) = RGB(0, 0, 196)
Colourz(3) = RGB(0, 196, 0)
Colourz(4) = RGB(255, 0, 0)
Colourz(5) = RGB(128, 0, 64)
Colourz(6) = RGB(196, 0, 196)
Colourz(7) = RGB(255, 128, 0)
Colourz(8) = RGB(255, 255, 0)
Colourz(9) = RGB(0, 255, 0)
Colourz(10) = RGB(0, 128, 128)
Colourz(11) = RGB(0, 255, 255)
Colourz(12) = RGB(0, 128, 255)
Colourz(13) = RGB(255, 0, 255)
Colourz(14) = RGB(128, 128, 128)
Colourz(15) = RGB(192, 192, 192)
Colourz(16) = RGB(255, 255, 255)

ColourzScheme(1) = "12"  ' Time
ColourzScheme(2) = "04"  ' Time Border
ColourzScheme(3) = "12"  ' Command
ColourzScheme(4) = "04"  ' Command Border
ColourzScheme(5) = "12"  ' Nickname
ColourzScheme(6) = "04"  ' Nickname Border
ColourzScheme(7) = "12"  ' Extended
ColourzScheme(8) = "04"  ' Extended Border
ColourzScheme(9) = "16"  ' Default Text
ColourzScheme(10) = "11" ' Default Text 2
ColourzScheme(11) = "00" ' Nothing Yet
ColourzScheme(12) = "00" ' Nothing Yet
ColourzScheme(13) = "00" ' Nothing Yet
ColourzScheme(14) = "00" ' Nothing Yet
ColourzScheme(15) = "00" ' Nothing Yet
ColourzScheme(16) = "00" ' Nothing Yet
End Sub

Function FadeColourz(ByVal Text As String) As String
  Dim Text2 As String
  Dim i As Integer, i2 As Byte
  Dim ColourList(6) As Byte
  ColourStyle = 2
  Select Case ColourStyle
    Case 1
      ColourList(0) = 4
      ColourList(1) = 6
      ColourList(2) = 12
      ColourList(3) = 11
      ColourList(4) = 9
      ColourList(5) = 8
      For i = 1 To Len(Text)
        If 1 = 2 Then
          i2 = i2 + 1
          If i2 = 16 Then i2 = 1
          If i2 < 10 Then
            Text2 = Text2 & "0" & i2 & Mid(Text, i, 1)
          Else
            Text2 = Text2 & "" & i2 & Mid(Text, i, 1)
          End If
        Else
          If i2 = 6 Then i2 = 0
'          If ColourList(i2) < 10 Then
'            Text2 = Text2 & "0" & ColourList(i2) & Mid(Text, i, 1)
'          Else
            Text2 = Text2 & "" & ColourList(i2) & Mid(Text, i, 1)
'          End If
          i2 = i2 + 1
        End If
      Next i
    
    Case 2
      Dim tmp1 As Integer, tmp2 As String
      Do
        tmp1 = InStr(Text, " ") - 1
        If tmp1 < 1 Then tmp1 = Len(Text)
        tmp2 = Mid(Text, 1, tmp1)
        
        If Len(tmp2) = 1 Then
          Text2 = Text2 & "3" & tmp2 & " "
        ElseIf Len(tmp2) = 2 Then
          Text2 = Text2 & "6" & tmp2 & " "
        Else
          Text2 = Text2 & "6" & Left(tmp2, 1) & "3" & Mid(tmp2, 2, Len(tmp2) - 2) & "6" & Right(tmp2, 1) & " "
        End If
        Text = Mid(Text, tmp1 + 2)
      Loop Until Len(Text) = 0
    Case Else
      Text2 = Text
  End Select
  FadeColourz = Trim(Text2)
End Function
