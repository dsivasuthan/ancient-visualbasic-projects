VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Speak it - Dhayalan Sivasuthan"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
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
   ScaleHeight     =   5025
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Set"
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ComboBox Speed 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Form1.frx":0ECA
      Left            =   120
      List            =   "Form1.frx":0EE9
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   615
      Left            =   600
      OleObjectBlob   =   "Form1.frx":0F1A
      TabIndex        =   6
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Speaking"
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0F3E
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Speak it"
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":0F67
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":14F1
      Top             =   120
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   5520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan's"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Speak It !"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
TextToSpeech1.Speak Text1.Text
End Sub

Private Sub Command2_Click()
TextToSpeech1.StopSpeaking
End Sub

Private Sub Command3_Click()
Me.Enabled = False
Dialog.Show
End Sub

Private Sub Command4_Click()
TextToSpeech1.StopSpeaking
TextToSpeech1.Speed = Speed.Text
End Sub

