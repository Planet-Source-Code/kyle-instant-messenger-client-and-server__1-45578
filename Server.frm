VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Server 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   3570
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   2175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   2175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMessage 
      Caption         =   "Message Users"
      Height          =   615
      Left            =   720
      Picture         =   "Server.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdKick 
      Caption         =   "Kick"
      Height          =   615
      Left            =   120
      Picture         =   "Server.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3315
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Status : Server Running"
            TextSave        =   "Status : Server Running"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Message to send :"
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtMess 
         Appearance      =   0  'Flat
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send Message"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.Frame fraLoggedOn 
      Appearance      =   0  'Flat
      Caption         =   "User's Online : "
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      Begin VB.ListBox lstUsers 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Timer tmrConnect 
      Interval        =   10000
      Left            =   2520
      Top             =   3360
   End
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   3360
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCloseServer 
         Caption         =   "&Close Server"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuExit11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "&Options"
      Begin VB.Menu mnuMessage 
         Caption         =   "Message All users"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuBreaker88 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKickUser 
         Caption         =   "&Kick User.."
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuDataTrans 
         Caption         =   "Data Transfer.."
      End
      Begin VB.Menu mnuLogged 
         Caption         =   "Logged Data"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, Y As Integer
Private Sub cmdKick_Click()
    Call mnuKickUser_Click
End Sub
Private Sub cmdMessage_Click()
    Call mnuMessage_Click
End Sub
Private Sub Command1_Click()
    For i% = 1 To WS().UBound
        Select Case WS(i%).State
            Case Is = sckConnected
                WS(i%).SendData "MESSAGE¤¶£" & Chr(198) & txtMess.Text
                Status "Sending - " & i%
                DoEvents%
            Case Else
        End Select
    Next i%
    AddData "MESSAGE", "MESSAGE¤¶£" & Chr(198) & txtMess.Text
    Status "Sent"
End Sub
Private Sub Form_Load()
    WS(0).Close
    WS(0).LocalPort = 12584
    WS(0).Listen
    Load Logg
    Load Data
    With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Cruel Instant Messenger Server - By Kyle" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nidProgramData
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
End Sub
Private Sub mnuCloseServer_Click()
    For i% = 1 To WS().UBound
        Select Case WS(i%).State
            Case Is = sckConnected
                WS(i%).Close
                Status "Closing - " & i%
                DoEvents%
        End Select
    Next i%
    AddData "Sock Close", "Ws" & (i%) & ".Close"
    Status "Server Closed"
    lstUsers.Clear
End Sub
Private Sub mnuDataTrans_Click()
    Data.Show
End Sub
Private Sub mnuExit_Click()
    Shell_NotifyIcon NIM_DELETE, nidProgramData
    End
End Sub
Private Sub mnuKickUser_Click()
    Select Case lstUsers.Text
        Case Is = ""
            Status ("Highlight User")
        Case Is <> ""
            For i% = 1 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData "BOOT¤¶£" & Chr(198) & lstUsers.Text
                        DoEvents%
                    Case Else
                End Select
            Next i%
            AddData "BOOT", "BOOT¤¶£" & Chr(198) & lstUsers.Text
    End Select
End Sub
Private Sub mnuLogged_Click()
    Logg.Show
End Sub
Private Sub mnuMessage_Click()
    If mnuMessage.Checked = False Then
        mnuMessage.Checked = True
        Me.Width = 4575
    ElseIf mnuMessage.Checked = True Then
        mnuMessage.Checked = False
        Me.Width = 2295
    End If
End Sub
Private Sub tmrConnect_Timer()
    Select Case WS(0).State
        Case Is <> sckConnected
            WS(0).Close
            WS(0).LocalPort = 12584
            WS(0).Listen
        Case Else
    End Select
End Sub
Private Sub WS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim NextSocket As Integer
    NextSocket% = WS().UBound + 1
    Load WS(NextSocket%)
    WS(NextSocket%).Accept (requestID)
End Sub
Private Sub WS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, Func As String, Dat As String
Dim From As String, Text As String
    Call WS(Index).GetData(Data$, vbString)
    Func$ = Split(Data$, "¤¶£" & Chr(198))(0)
    Dat$ = Split(Data$, "¤¶£" & Chr(198))(1)
    Select Case Func$
        Case Is = "LOGIN"
            lstUsers.AddItem Dat$
            For Y% = 1 To lstUsers.ListCount
                lstUsers.ListIndex = Y% - 1
                Select Case lstUsers.Text
                    Case Is <> Dat$
                        WS(WS().UBound).SendData "NICK¤¶£" & Chr(198) & lstUsers.Text
                        DoEvents%
                    Case Else
                End Select
            Next Y%
            For i% = 1 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData "NICK¤¶£" & Chr(198) & Dat$
                        DoEvents%
                        Pause 0.05
                        AddData "Username", Data$
                End Select
            Next i%
            AddData "Username", Data$
            Log Dat$, WS(Index).RemoteHostIP
        Case Is = "GETUSER"
            For i% = 0 To lstUsers.ListCount - 1
                lstUsers.ListIndex = (i%)
                WS(Index).SendData ("NICK¤¶£" & Chr(198) & lstUsers.List(i%))
                DoEvents%
            Next i%
            AddData "GetUser", Data$
        Case Is = "PM"
            For i% = 0 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData (Data$)
                        DoEvents%
                End Select
            Next i%
            AddData "Pm", Data$
        Case Is = "LOGOUT"
            Call RemoveItemFromListbox(lstUsers, Dat$)
            For i% = 1 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData (Data$)
                        DoEvents%
                    Case Else
                End Select
            Next i%
            AddData "Logout", Data$
        Case Is = "CHAT"
            For i% = 1 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData (Data$)
                        DoEvents%
                    Case Else
                End Select
            Next i%
            AddData "ChatSend", Data$
        Case Is = "ENTER"
            For i% = 1 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData (Data$)
                        DoEvents%
                    Case Else
                End Select
            Next i%
            AddData "ChatRoomEnter", Data$
        Case Is = "LEAVE"
            For i% = 1 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData (Data$)
                        DoEvents%
                    Case Else
                End Select
            Next i%
            AddData "ChatRoomLeave", Data$
        Case Is = "IGNORE"
            For i% = 1 To WS().UBound
                Select Case WS(i%).State
                    Case Is = sckConnected
                        WS(i%).SendData (Data$)
                        DoEvents%
                    Case Else
                End Select
            Next i%
            AddData "Ignored", Data$
    End Select
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo Form_MouseMove_err:
    Dim Result, MSG         As Long
    If Me.ScaleMode = vbPixels Then
        MSG = x
    Else
        MSG = x / Screen.TwipsPerPixelX
    End If
    Select Case MSG
        Case WM_LBUTTONUP
            Me.Show
            Me.WindowState = 0
        Case WM_LBUTTONDBLCLK
            Me.Show
        Case WM_RBUTTONUP
            PopupMenu Server.mnuFile
    End Select
Form_MouseMove_err:
End Sub


