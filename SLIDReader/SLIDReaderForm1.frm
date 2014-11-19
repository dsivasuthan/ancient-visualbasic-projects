VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SL Identity Reader"
   ClientHeight    =   2175
   ClientLeft      =   1335
   ClientTop       =   465
   ClientWidth     =   3945
   Icon            =   "SLIDReaderForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWidth 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   2640
   End
   Begin VB.Timer tmrHeight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   2640
   End
   Begin VB.Timer tmrBlue 
      Interval        =   2000
      Left            =   3240
      Top             =   3120
   End
   Begin VB.Timer tmrGreen 
      Interval        =   1000
      Left            =   2880
      Top             =   3120
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MouseIcon       =   "SLIDReaderForm1.frx":1CCA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "000000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1097
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      MaxLength       =   9
      MouseIcon       =   "SLIDReaderForm1.frx":A194
      TabIndex        =   1
      ToolTipText     =   "Enter Your ID card No"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   2880
      Picture         =   "SLIDReaderForm1.frx":1265E
      Top             =   -135
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2085
      TabIndex        =   10
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Gender"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Your Birth year"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your ID no and press Search"
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
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblGender 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblYear 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If txtID = "" Then
    MsgBox "Enter ID card no", , "Siva"
Else
    lblYear = Val(Left(txtID.Text, 2)) + 1900
    If Val(Mid(txtID.Text, 3, 3)) > 500 Then
    lblGender.Caption = "Female"
Else
    lblGender.Caption = "Male"
    End If
End If
End Sub

Private Sub Command2_Click()

Dim Ex As Integer
Ex = MsgBox("Do you really want to exit?", vbYesNo, "Siva")
If Ex = vbYes Then
tmrHeight.Enabled = True
End If
End Sub

Private Sub Command3_Click()
txtID = ""
lblGender = ""
lblYear = ""
End Sub

Private Sub Command4_Click()
Me.Enabled = False
Dialog.Show
End Sub

Private Sub Form_Load()
MsgBox "Welcome to Siva's Identity Card Reader", vbOKOnly, "Siva"
End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Dim Ex As Integer
Ex = MsgBox("Do you really want to exit?", vbYesNo, "Siva")
If Ex = vbYes Then
tmrHeight.Enabled = True
End If
End Sub

Private Sub tmrBlue_Timer()
Label5.ForeColor = vbBlue
End Sub

Private Sub tmrGreen_Timer()
Label5.ForeColor = vbGreen
End Sub

Private Sub tmrHeight_Timer()
If Me.Height > 600 Then
Me.Height = Me.Height - 30
Else
End
End If
End Sub

'Private Sub tmrWidth_Timer()
'If Me.Width > 2500 Then
'Me.Width = Me.Width - 30
'Else
'End
'End If
'End Sub

'Private Sub txtID_KeyPress(KeyAscii As Integer)
''If KeyAscii > 70 Then
'MsgBox "Error"
'txtID.Text = Val((txtID.Text, 3, 8)) + 1
'End If
'End Sub
Private Sub txtID_Change()

End Sub

Private Sub txtID_GotFocus()
Command1.Default = True

End Sub
