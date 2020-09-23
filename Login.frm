VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Login :"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtHost 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "24.116.47.27"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Host :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim D As Integer, L As Integer, i As Integer
Private Sub cmdLogin_Click()
    Select Case txtUser.Text
        Case Is = ""
            NewError "You must put in a username to logon Punk Rawk Messenger."
        Case Else
            WS.Close
            WS.Connect txtHost.Text, 12584
    End Select
End Sub
Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = 13
            Call cmdLogin_Click
            KeyCode = 0
    End Select
End Sub
Private Sub WS_Connect()
    'when connected, it sends the Login string to the server
    Send "LOGIN¤¶£" & Chr(198) & txtUser.Text
    Me.Visible = False
    Status "Logged In"
    Mess.lstUsers.Nodes.Add , , "MAIN", "Online - " & Login.txtUser.Text, 2, 2
    Mess.lstUsers.Nodes.item(1).Expanded = True
    Mess.Caption = txtUser.Text
    Mess.mnuLogin.Enabled = False
    Mess.mnuClose.Enabled = True
End Sub
Private Sub WS_DataArrival(ByVal bytesTotal As Long)
'Error Control
'---------------------
On Error Resume Next
'Declare our Arrays
'-----------------------------------------------------------------
Dim Data As String, Func As String, dat As String
Dim Too As String, From As String, Name As String, Text As String
Dim X As Integer
    'Get Data, Store data in Array 'Data$'
    Call WS.GetData(Data$, vbString)
    'Split Data, Func$ is our Function that we wish to do, Dat$ is our Text
    'on the second side of our Seperator, I had to use an Ascii seperator
    'so people couldn't mess with the strings.
    '-----------------------------------------
    Func$ = Split(Data$, "¤¶£" & Chr(198))(0)
    dat$ = Split(Data$, "¤¶£" & Chr(198))(1)
    Select Case Func$
        'Add's User's name to the User list
        Case Is = "NICK"
            Mess.lstUsers.Nodes.Add "MAIN", tvwChild, , (dat$), 1, 1
            Pause 0.06
            Status (Mess.lstUsers.Nodes.Count - 1) & " user(s)"
        'PM From a user
        Case Is = "PM"
            'Split PM String, Too$ is who the Pm is from, From$ is the text the user
            'has sent in the string
            Too$ = Split(Data$, "¤¶£" & Chr(198))(2)
            From$ = Split(Data$, "¤¶£" & Chr(198))(3)
            'If the pm is meant for the user, a Pm box will show
            If dat$ = Login.txtUser.Text Then
                If InStr(1, Mess.txtIggy.Text, Too$) Then
                    Send "IGNORE" & "¤¶£" & Chr(198) & dat$ & "¤¶£" & Chr(198) & Too$
                Else
                    'Boo did part of it, But he messed a portion up, so i fixed it
                    'if the pm is already open then it will just add it to the textbox
                    For X% = 0 To 30
                        If NewBoo(X%).txtTo.Text = Too$ Then
                            NewBoo(X%).txtTo.Text = Too$
                            NewBoo(X%).txtFrom.Text = (Login.txtUser.Text)
                            NewBoo(X%).txtChat.SelStart = Len(NewBoo(X%).txtChat.Text)
                            NewBoo(X%).txtChat.SelBold = True
                            NewBoo(X%).txtChat.SelColor = vbBlue
                            NewBoo(X%).txtChat.SelText = Too$ & " : "
                            NewBoo(X%).txtChat.SelBold = False
                            NewBoo(X%).txtChat.SelColor = vbBlack
                            NewBoo(X%).txtChat.SelText = From$ & vbCrLf
                            NewBoo(X%).txtTo.Locked = True
                            NewBoo(X%).stat.Panels.item(1).Text = "Last Message Received on " & Date & " at " & Time
                            Exit Sub
                        End If
                    Next X%
                    'if the pm isn't open, then it opens a new one
                    NewBoo(D%).txtTo.Text = (Too$)
                    NewBoo(D%).txtFrom.Text = (Login.txtUser.Text)
                    NewBoo(D%).txtChat.SelStart = Len(NewBoo(D%).txtChat.Text)
                    NewBoo(D%).txtChat.SelBold = True
                    NewBoo(D%).txtChat.SelColor = vbBlue
                    NewBoo(D%).txtChat.SelText = Too$ & " : "
                    NewBoo(D%).txtChat.SelBold = False
                    NewBoo(D%).txtChat.SelColor = vbBlack
                    NewBoo(D%).txtChat.SelText = From$ & vbCrLf
                    NewBoo(D%).Caption = "Private Message -- " & NewBoo(D%).txtTo.Text
                    NewBoo(D%).stat.Panels.item(1).Text = "Last Message Received on " & Date & " at " & Time
                    NewBoo(D%).txtTo.Locked = True
                    NewBoo(D%).Show
                    'also add's to the d% integer, so next time it opens up
                    'a new pm
                    D% = D% + 1
                End If
            End If
        Case Is = "IGNORE"
            From$ = Split(Data$, "¤¶£" & Chr(198))(2)
            If From$ = Login.txtUser.Text Then
                For i% = 0 To 30
                    If NewBoo(i%).Caption = "Private Message -- " & dat$ Then
                        With NewBoo(i%)
                            .txtChat.SelStart = Len(.txtChat.Text)
                            .txtChat.SelBold = True
                            .txtChat.SelColor = vbBlack
                            .txtChat.SelText = "<User has ignored you>" & vbCrLf
                            Exit Sub
                        End With
                    End If
                Next i%
            Else
            End If
        'shows that a user has logged out of the server
        Case Is = "LOGOUT"
            Status "User(s) Left"
        'if func$ = message then the admin box will show up saying what the admin
        'wants you to know
        Case Is = "MESSAGE"
            Admin.txtMess.Text = (dat$)
            Admin.lblTime.Caption = "Time : " & Time
            Admin.lblDate.Caption = "Date : " & Date
            Admin.Show
        'if chat string is received, it puts the data into the Chat txtChat.text
        Case Is = "CHAT"
            'Splits data, splits User and Text
            Text$ = Split(Data$, "¤¶£" & Chr(198))(2)
            If dat$ = Login.txtUser.Text Then
            Else
                Chat.txtChat.SelStart = Len(Chat.txtChat.Text)
                Chat.txtChat.SelColor = vbBlue
                Chat.txtChat.SelBold = True
                Chat.txtChat.SelText = dat$ & " : "
                Chat.txtChat.SelBold = False
                Chat.txtChat.SelColor = vbBlack
                Chat.txtChat.SelText = Text$ & vbCrLf
            End If
        'if the admin has sent the boot string, it notifies the user that they have
        'been booted
        Case Is = "BOOT"
            If dat$ = Login.txtUser.Text Then
                MsgBox "Punk Rawk Messenger Server Admin has Kicked you off of the server.", vbInformation, "Disconnection"
                Mess.lstUsers.Nodes.Clear
                Send "LOGOUT¤¶£" & Chr(198) & Login.txtUser.Text
                Status "Logged Out"
                Mess.mnuLogin.Enabled = True
                Mess.mnuClose.Enabled = False
                txtUser.Text = ""
                Mess.Caption = "Punk Rawk Messenger"
                Pause (1)
                WS.Close
            End If
        'a user has entered chat
        Case Is = "ENTER"
            With Chat
                .txtChat.SelStart = Len(.txtChat.Text)
                .txtChat.SelBold = False
                .txtChat.SelColor = vbRed
                .txtChat.SelText = "<" & dat$ & " has joined the room>" & vbCrLf
            End With
        'a user has left the chat room
        Case Is = "LEAVE"
            With Chat
                .txtChat.SelStart = Len(.txtChat.Text)
                .txtChat.SelBold = False
                .txtChat.SelColor = vbRed
                .txtChat.SelText = "<" & dat$ & " has left the room>" & vbCrLf
            End With
    End Select
End Sub
Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'if there is an error, it will display Connection Error
    Status "Connection Error"
End Sub
