VERSION 5.00
Begin VB.Form frmWelcomescreen 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome - Dhayalan Sivasuthan"
   ClientHeight    =   3945
   ClientLeft      =   5505
   ClientTop       =   2985
   ClientWidth     =   6375
   Icon            =   "Welcomescreen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   600
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   5640
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6120
      Top             =   480
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Chemical Equations"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   6135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Atomicity"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Valency"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Elements && Their Symbols"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "By Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Click on one of the topics to continue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Explore Chemistry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   630
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmWelcomescreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
frmSymbols.Show
End Sub

Private Sub Command2_Click()
Me.Hide
frmValency.Show
End Sub

Private Sub Command3_Click()
Me.Hide
frmAtomicity.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Command6_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Form_Load()
Label2.FontUnderline = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Dim Ex As String
Ex = MsgBox("Do you really want to quit Siva's Chemistry quiz?", vbYesNo, "siva")
If Ex = vbYes Then
Timer3.Enabled = True
Dialog.Show
End If
End Sub

Private Sub Label3_Click()
Command6_Click
End Sub

Private Sub Timer1_Timer()
Label3.ForeColor = vbBlue
Label1.ForeColor = vbGreen
End Sub

Private Sub Timer2_Timer()
Label3.ForeColor = vbGreen
Label1.ForeColor = vbRed
End Sub

Private Sub Timer3_Timer()
If Me.Height > 700 Then
Me.WindowState = Normal
Me.Height = Me.Height - 50

Else
End
End If
End Sub
