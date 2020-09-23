VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtINfo 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.ListBox lstError 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "Help.frx":0CCA
      Left            =   120
      List            =   "Help.frx":0CE9
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin MSComctlLib.StatusBar Stat 
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
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
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Info. on Function :"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Functions Used :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Above is the status displayer found on the Messenger Client."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstError_Click()
    Select Case lstError.ListIndex
        Case Is = 0
            Call Hel
            txtINfo.Text = "Connection Error : " & vbCrLf & "This function is displayed when the Server is not running, so you cannot get connected to the server."
        Case Is = 1
            Call Hel
            txtINfo.Text = "New User(s) : " & vbCrLf & "This is displayed when a new user has logged onto the client, when displayed, it would be wise to click Get Online Users, to 'Refresh' your list of Online Users."
        Case Is = 2
            Call Hel
            txtINfo.Text = "Logged In : " & vbCrLf & "This is displayed when you are Successfully logged onto the Instant messenger server, after logged in, you should push the 'Get Online Users' button to see who is online."
        Case Is = 3
            Call Hel
            txtINfo.Text = "Logged Out : " & vbCrLf & "This meens you are successfully logged out of the chat server, and if you wish to chat again, you have to Re-Log back in."
        Case Is = 4
            Call Hel
            txtINfo.Text = "Highlight User : " & vbCrLf & "This is displayed when you have not highlighted a user to 'PM' or 'Ignore', you must click on a name in the list and push 'PM' or 'Iggy'."
        Case Is = 5
            Call Hel
            txtINfo.Text = "Not Connected : " & vbCrLf & "This is displayed when the Client is trying to send data, but it cannot because you are not connected.If this shows up, it would be wise to Logout, then Login again."
        Case Is = 6
            Call Hel
            txtINfo.Text = "Private Message : " & vbCrLf & "'PM' stands for Private Message, this program was basically only made for Private Messaging, so that's what it is used for."
        Case Is = 7
            Call Hel
            txtINfo.Text = "Iggy User : " & vbCrLf & "Iggy is used for if you want to 'Ignore' someone, just simply click on their name, and hit the Iggy button, and you will no longer receive message's from the user."
        Case Is = 8
            Call Hel
            txtINfo.Text = "User(s) Left : " & vbCrLf & "This is displayed when a user has logged out of the server, if you want to see who logged out, just click 'Get Online Users' and see who logged off."
    End Select
End Sub
Sub Hel()
    Me.Caption = "Help - " & lstError.Text
End Sub
