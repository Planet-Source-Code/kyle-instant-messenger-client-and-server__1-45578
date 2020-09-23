VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Message --"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"PM.frx":038A
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   3135
   End
End
Attribute VB_Name = "PM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    Select Case txtSend.Text
        Case Is = ""
        Case Is <> ""
            Send "PM" & "¤¶£" & Chr(198) & txtTo.Text & "¤¶£" & Chr(198) & txtFrom.Text & "¤¶£" & Chr(198) & txtSend.Text
            Debug.Print "PM" & "¤¶£" & Chr(198) & txtTo.Text & "¤¶£" & Chr(198) & txtFrom.Text & "¤¶£" & Chr(198) & txtSend.Text
            txtChat.SelStart = Len(txtChat.Text)
            txtChat.SelColor = vbRed
            txtChat.SelBold = True
            txtChat.SelText = txtFrom.Text & " : "
            txtChat.SelColor = vbBlack
            txtChat.SelBold = False
            txtChat.SelText = txtSend.Text & vbCrLf
            txtTo.Locked = True
            txtSend.Text = ""
    End Select
End Sub
Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = 13
            Call cmdSend_Click
            KeyCode = 0
    End Select
End Sub
