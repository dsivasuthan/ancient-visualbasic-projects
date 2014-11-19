VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About the Author - Dhayalan Sivasuthan"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8250
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   8055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   600
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Interval        =   3000
      Left            =   1080
      Top             =   1200
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   7440
      Picture         =   "Form2.frx":1CCA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "Form2.frx":13A5A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label17 
      Caption         =   "Negombo South International School, Negombo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label16 
      Caption         =   "16 years"
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
      Left            =   5160
      TabIndex        =   16
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label15 
      Caption         =   "www.dsiva.8m.com"
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
      Left            =   5160
      TabIndex        =   15
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label14 
      Caption         =   "dhayalansivasuthan@yahoo.com"
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
      Left            =   5160
      TabIndex        =   14
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label13 
      Caption         =   "0786109828 / 0779706600"
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
      Left            =   5160
      TabIndex        =   13
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "#28,Wendeese watte, Karambe, Palavi, Puttalam, Sri Lanka"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label11 
      Caption         =   "Dhayalan Sivasuthan"
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
      Left            =   5160
      TabIndex        =   11
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "Address"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "School:"
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
      Left            =   4200
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Age:"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Website:"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Email:"
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
      Left            =   4200
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Tel No:"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Address:"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   3015
      Left            =   120
      Top             =   840
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   240
      Picture         =   "Form2.frx":15724
      Stretch         =   -1  'True
      Top             =   960
      Width           =   3600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "dhayalansivasuthan@yahoo.com"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   7440
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form1.Enabled = True
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Form1.Enabled = True
Me.Hide
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = vbGreen
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = vbBlue
End Sub

Private Sub Timer3_Timer()
Label1.ForeColor = vbRed
End Sub
