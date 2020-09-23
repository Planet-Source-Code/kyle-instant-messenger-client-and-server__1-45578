VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About - Creator"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":038A
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":0429
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location : Fargo, North Dakota"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label lblCreator 
      BackStyle       =   0  'Transparent
      Caption         =   "Creator : Kyle Wald"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Image imgMe 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   120
      Picture         =   "About.frx":04B8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2610
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
