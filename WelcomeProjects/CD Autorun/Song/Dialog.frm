VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dhayalan Sivasuthan"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "www.dsiva.8m.com"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "dhayalansivasuthan@yahoo.com"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "078 6109828 / 077 9706600"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "#28, Wendeese watte, Karambe, Palavi, Puttalam, Sri Lanka"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   8
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "16 years old"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Web"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel No"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1410
      Left            =   120
      Picture         =   "Dialog.frx":57E2
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
If DEx = True Then
End
Else
Me.Hide
Form1.Show
Form1.Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If DEx = True Then
End
Else

Me.Hide
Form1.Show
Form1.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Label1(1).ForeColor = RGB(Rnd * 500, Rnd * 500, Rnd * 500)
End Sub
