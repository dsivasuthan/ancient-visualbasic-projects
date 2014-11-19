VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmWelcomeScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Converter - D.Sivasuthan"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "WelcomeScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   3000
      Left            =   4080
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   3600
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   2880
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   4215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "About Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Temperature"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Energy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Pressure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Mass"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Force"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Density"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This was made by Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the button for conversion:"
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
      TabIndex        =   13
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmWelcomeScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IntRes As Integer
Private Sub Command1_Click()
Me.Hide
frmLengthConversion.Show
End Sub

Private Sub Command10_Click()
Me.Hide
frmEnergyConversion.Show
End Sub

Private Sub Command11_Click()
Me.Hide
frmTemperatureConversion.Show
End Sub

Private Sub Command12_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Command13_Click()

'Dim Ex As String
'Ex = MsgBox("Do you want quit Siva's Convertor?", vbYesNo, "Siva")
'If Ex = vbYes Then
Timer4.Enabled = True
Dialog.Show
'End If
End Sub

Private Sub Command2_Click()
Me.Hide
frmAreaConversion.Show
End Sub

Private Sub Command3_Click()
Me.Hide
frmSpeedConversion.Show
End Sub

Private Sub Command4_Click()
Me.Hide
frmVolumeConversion.Show
End Sub

Private Sub Command5_Click()
frmDensityConversion.Show
Me.Hide
End Sub

Private Sub Command7_Click()
frmMassConversion.Show
Me.Hide
End Sub



Private Sub Command8_Click()
Me.Hide
frmPressureConversion.Show
End Sub

Private Sub Command9_Click()
Me.Hide
frmPowerConversion.Show
End Sub

Private Sub Form_Load()
'IntRes = MsgBox("Welcome to Siva's Converter!", vbOKOnly, "Welcome")
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim Ex As String
Cancel = True

'Ex = MsgBox("Do you want quit Siva's Convertor?", vbYesNo, "Siva")
'If Ex = vbYes Then
Timer4.Enabled = True
Dialog.Show
'End If
End Sub

Private Sub Timer1_Timer()
Label2.ForeColor = vbGreen
'Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Label2.ForeColor = vbRed

'Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Label2.ForeColor = vbBlue
End Sub

Private Sub Timer4_Timer()
If Me.Height > 700 Then
Me.Height = Me.Height - 50
Else
End
End If
End Sub
