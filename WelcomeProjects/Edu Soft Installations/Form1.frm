VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dhayalan Sivasuthan - Education"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "SOLVE ELECTRIC"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4800
      Width           =   4815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "STANDING WAVE"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   4815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NEWTON's MAZE"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   4815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FREEFALL"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DATAFIT"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AIRTRACKS"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AIRTABLE"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RELATIVITY"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click on the following buttons to install the corrsponding software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Educational Software Intallations"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "ACB2004_RELATIVITY.exe"
End Sub

Private Sub Command2_Click()
Shell "AIRTABLE.exe"

End Sub

Private Sub Command3_Click()
Shell "AIRTRACKS.exe"
End Sub

Private Sub Command4_Click()
Shell "DATAFIT2000.exe"
End Sub

Private Sub Command5_Click()
Shell "FREEFALL.exe"
End Sub

Private Sub Command6_Click()
Shell "NEWTONS_MAZE.exe"
End Sub

Private Sub Command7_Click()
Shell "STANDINGWAVE.exe"
End Sub

Private Sub Command8_Click()
Shell "solveelec20ensetup.exe"
End Sub
