VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Chat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punk Rawk Messenger Chat"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Chat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   1
      Left            =   2160
      Top             =   1920
   End
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   4095
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Chat.frx":038A
   End
End
Attribute VB_Name = "Chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    Select Case txtSend.Text
        Case Is = ""
        Case Is <> ""
            Send "CHAT" & "¤¶£" & Chr(198) & Login.txtUser.Text & "¤¶£" & Chr(198) & txtSend.Text
            Debug.Print "CHAT" & "¤¶£" & Chr(198) & Login.txtUser.Text & "¤¶£" & Chr(198) & txtSend.Text
            txtChat.SelStart = Len(txtChat.Text)
            txtChat.SelColor = vbRed
            txtChat.SelBold = True
            txtChat.SelText = Login.txtUser.Text & " : "
            txtChat.SelColor = vbBlack
            txtChat.SelBold = False
            txtChat.SelText = txtSend.Text & vbCrLf
            txtSend.Text = ""
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Send "LEAVE" & "¤¶£" & Chr(198) & Login.txtUser.Text
    Me.Visible = False
End Sub
Private Sub tmrTime_Timer()
    stat.Panels.item(1).Text = Time
    stat.Panels.item(2).Text = Date
End Sub
Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = 13
            Call cmdSend_Click
            KeyCode = 0
    End Select
End Sub
