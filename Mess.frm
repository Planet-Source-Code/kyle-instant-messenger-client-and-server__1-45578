VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Mess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punk Rawk Messenger"
   ClientHeight    =   4110
   ClientLeft      =   1935
   ClientTop       =   3945
   ClientWidth     =   2640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   2640
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Online"
      TabPicture(0)   =   "Mess.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstUsers"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdChat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdIggy"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Iggy List"
      TabPicture(1)   =   "Mess.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIggy"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox txtIggy 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   5318
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Mess.frx":03C2
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh List"
         Height          =   255
         Left            =   120
         Picture         =   "Mess.frx":043D
         TabIndex        =   7
         ToolTipText     =   "Get "
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton cmdPm 
         Caption         =   "PM"
         Height          =   615
         Left            =   120
         Picture         =   "Mess.frx":07C7
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Private Message User"
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton cmdIggy 
         Caption         =   "Iggy"
         Height          =   615
         Left            =   720
         Picture         =   "Mess.frx":0B51
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ignore User"
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdChat 
         Caption         =   "Chat"
         Height          =   615
         Left            =   1680
         Picture         =   "Mess.frx":0EDB
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   495
      End
      Begin MSComctlLib.TreeView lstUsers 
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "IL"
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      Picture         =   "Mess.frx":1265
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   2400
      Width           =   255
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   2760
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mess.frx":15EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mess.frx":1911
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mess.frx":1B7F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3855
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Status : Idle"
            TextSave        =   "Status : Idle"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogin 
         Caption         =   "Login"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Logout"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBreaker2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuIggy 
         Caption         =   "Iggy User"
      End
      Begin VB.Menu mnuUnIggy 
         Caption         =   "Un-Iggy Users"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBreaker989 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChat 
         Caption         =   "Chat.."
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuNeedhelp 
         Caption         =   "&Need Help?"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About.."
      End
   End
End
Attribute VB_Name = "Mess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Instant Messanger : By Kyle W.
Option Explicit
Dim D As Integer, X As Integer, i As Integer
Private Sub cmdChat_Click()
    Select Case Login.WS.State
        Case Is = sckConnected
            If Chat.Visible = True Then
            Else
                Call mnuChat_Click
                Send "ENTER" & "¤¶£" & Chr(198) & Login.txtUser.Text
            End If
        Case Else
            Status ("Not Connected")
    End Select
End Sub
Private Sub cmdIggy_Click()
On Error Resume Next
Dim Ignore As String
    Select Case lstUsers.SelectedItem.Text
        Case Is = Login.txtUser.Text
            NewError "You cannot ignore yourself."
        Case Else
            If cmdIggy.Caption = "Un-Iggy" Then
                Ignore$ = Split(lstUsers.SelectedItem.Text, " (")(0)
                lstUsers.SelectedItem.Text = (Ignore$)
                txtIggy.Find Ignore$ & vbCrLf
                txtIggy.SelText = ""
                cmdIggy.Caption = "Iggy"
            ElseIf cmdIggy.Caption = "Iggy" Then
                Call mnuIggy_Click
                cmdIggy.Caption = "Un-Iggy"
            End If
    End Select
End Sub
Private Sub Command1_Click()
    Select Case Login.WS.State
        Case Is = sckConnected
            lstUsers.Nodes.Clear
            Send "GETUSER¤¶£" & Chr(198)
            Mess.lstUsers.Nodes.Add , , "MAIN", "Online - " & Login.txtUser.Text, 2, 2
            Mess.lstUsers.Nodes.item(1).Expanded = True
        Case Else
            Status ("Not Connected")
    End Select
End Sub
Private Sub cmdpm_Click()
On Error Resume Next
Dim Ignore As String
    Select Case lstUsers.SelectedItem.Text
        Case Is = ""
            Status "Highlight User"
        Case Is <> ""
            Select Case lstUsers.SelectedItem.Text
                Case Is = Login.txtUser.Text
                    NewError "You cannot Pm yourself."
                Case Is = "Online - " & Login.txtUser.Text
                    Exit Sub
                Case Else
                    If InStr(1, lstUsers.SelectedItem, "Ignored") Then
                        Ignore$ = Split(lstUsers.SelectedItem.Text, " (")(0)
                        NewError "Cannot send an instant message to " & Ignore$ & " because you have ignored the user"
                    Else
                        For X% = 0 To 20
                            If NewBoo(X%).Caption = "Private Message -- " & lstUsers.SelectedItem.Text Then
                                Exit Sub
                            End If
                        Next X%
                        NewBoo(D%).txtTo.Text = lstUsers.SelectedItem.Text
                        NewBoo(D%).txtFrom.Text = Login.txtUser.Text
                        NewBoo(D%).Caption = "Private Message -- " & NewBoo(D%).txtTo.Text
                        NewBoo(D%).Show
                        D% = D% + 1
                    End If
            End Select
    End Select
End Sub
Private Sub mnuClone_Click()
    Dim Login As New Login
    Login.Show
End Sub
Private Sub mnuCloneClient_Click()
    Dim Mess As New Mess
    Mess.Show
End Sub
Private Sub Command2_Click()
    txtIggy.Text = ""
End Sub
Private Sub Form_Load()
    With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Punk Rawk Messenger - By Kyle" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nidProgramData
End Sub
Private Sub lstUsers_Click()
On Error Resume Next
    If lstUsers.SelectedItem.Index = 1 Then
        If lstUsers.SelectedItem.Expanded = True Then
            lstUsers.SelectedItem.Expanded = False
            lstUsers.SelectedItem.Image = 3
        ElseIf lstUsers.SelectedItem.Expanded = False Then
            lstUsers.SelectedItem.Expanded = True
            lstUsers.SelectedItem.Image = 2
        End If
    ElseIf InStr(1, lstUsers.SelectedItem.Text, "Ignored") Then
        cmdIggy.Caption = "Un-Iggy"
    Else
        cmdIggy.Caption = "Iggy"
    End If
End Sub
Private Sub lstUsers_DblClick()
    Call cmdpm_Click
End Sub
Private Sub mnuAbout_Click()
    About.Show
End Sub
Private Sub mnuChat_Click()
    Select Case Login.WS.State
        Case Is = sckConnected
            Chat.Show
        Case Else
            Status ("Not Connected")
    End Select
End Sub
Private Sub mnuClose_Click()
    Send "LOGOUT¤¶£" & Chr(198) & Login.txtUser.Text
    Status "Logged Out"
    mnuLogin.Enabled = True
    mnuClose.Enabled = False
    lstUsers.Nodes.Clear
    Me.Caption = "Punk Rawk Messenger"
    Login.txtUser.Text = ""
    Pause (1)
    Login.WS.Close
End Sub
Private Sub mnuExit_Click()
    Shell_NotifyIcon NIM_DELETE, nidProgramData
    Call mnuClose_Click
    Status ("Closing")
    End
End Sub
Private Sub mnuIggy_Click()
On Error Resume Next
    Select Case lstUsers.SelectedItem.Text
        Case Is = ""
            Status "Highlight User"
        Case Is <> ""
            Select Case lstUsers.SelectedItem
                Case Is = Login.txtUser.Text
                    NewError "You cannot ignore yourself."
                Case Else
                    txtIggy.Text = txtIggy.Text & lstUsers.SelectedItem & vbCrLf
                    lstUsers.SelectedItem.Text = lstUsers.SelectedItem.Text & " (Ignored)"
            End Select
    End Select
End Sub
Private Sub mnuLogin_Click()
    Login.Show
End Sub
Private Sub mnuNeedhelp_Click()
    Help.Show
End Sub
Private Sub mnuUnIggy_Click()
    txtIggy.Text = ""
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Form_MouseMove_err:
    Dim Result, MSG         As Long
    If Me.ScaleMode = vbPixels Then
        MSG = X
    Else
        MSG = X / Screen.TwipsPerPixelX
    End If
    Select Case MSG
        Case WM_LBUTTONUP
            Me.Show
            Me.WindowState = 0
        Case WM_LBUTTONDBLCLK
            Me.Show
        Case WM_RBUTTONUP
            PopupMenu Mess.mnuFile
    End Select
Form_MouseMove_err:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
End Sub

