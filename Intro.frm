VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Intro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prog 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Image imgBack 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2190
      Left            =   0
      Picture         =   "Intro.frx":0000
      Top             =   0
      Width           =   4350
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
    Me.Visible = True
    For i% = 0 To 100
        prog.Value = (i%)
        Pause (0.02)
        If (i% = 100) Then
            Mess.Show
            Unload Me
        End If
    Next i%
End Sub
