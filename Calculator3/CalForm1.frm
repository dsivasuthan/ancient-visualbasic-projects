VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Calculator - Dhayalan Sivasuthan"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   4275
   Icon            =   "CalForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2400
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Pie"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdSqr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Square"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2595
      Width           =   1095
   End
   Begin VB.CommandButton cmdLog 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2220
      Width           =   1095
   End
   Begin VB.CommandButton cmdTan 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1815
      Width           =   1095
   End
   Begin VB.CommandButton cmdSin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1410
      Width           =   1095
   End
   Begin VB.CommandButton cmdCos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   990
      Width           =   1095
   End
   Begin VB.CommandButton cmdSqrRt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Square Root"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   585
      Width           =   1095
   End
   Begin VB.CommandButton Dot 
      BackColor       =   &H000080FF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Equals 
      BackColor       =   &H00008000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "-/+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Div 
      BackColor       =   &H000000FF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Times 
      BackColor       =   &H000000FF&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Minus 
      BackColor       =   &H000000FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Plus 
      BackColor       =   &H000000FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Clear 
      BackColor       =   &H0080C0FF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H000080FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   26
      Top             =   2970
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Menu Cal 
      Caption         =   "Cal"
      Visible         =   0   'False
      Begin VB.Menu mode 
         Caption         =   "Mode"
         Begin VB.Menu Sci 
            Caption         =   "Scientific"
         End
         Begin VB.Menu Normal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu Spac 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu C 
         Caption         =   "Clear"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu Siva 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Operand1 As Double
Dim Operand2 As Double
Dim Operator As String
Dim ClearDisplay As Boolean

Private Sub C_Click()
Clear_Click
End Sub

Private Sub Clear_Click()
Display.Caption = ""
End Sub



Private Sub Close_Click()
End
End Sub

Private Sub cmdCos_Click()
On Error Resume Next

Display.Caption = Cos(Display.Caption)

End Sub

Private Sub cmdLog_Click()
On Error Resume Next

Display.Caption = Log(Display.Caption)

End Sub

Private Sub cmdSin_Click()
On Error Resume Next

Display.Caption = Sin(Display.Caption)

End Sub

Private Sub cmdSqr_Click()
On Error Resume Next

Display.Caption = Display.Caption * Display.Caption

End Sub

Private Sub cmdSqrRt_Click()
On Error Resume Next
If Val(Display.Caption) < 0 Then
MsgBox "Can't calculate square root of a neagtive number", , "Siva"
Else
 Display.Caption = Sqr(Display.Caption)
End If
End Sub

Private Sub cmdTan_Click()
On Error Resume Next

Display.Caption = Tan(Display.Caption)

End Sub

Private Sub Command1_Click()
Display.Caption = "3.142"
End Sub



Private Sub Command2_Click()
Clipboard.SetText Display.Caption

End Sub

Private Sub Command5_Click()
Display.Caption = -Val(Display.Caption)
End Sub

Private Sub Copy_Click()
Clipboard.SetText Display.Caption

End Sub

Private Sub Digits_Click(Index As Integer)
Display.Caption = Display.Caption + Digits(Index).Caption
End Sub

Private Sub Div_Click()
Operand1 = Val(Display.Caption)
Operator = "/"
Display.Caption = ""
End Sub

Private Sub Dot_Click()
If InStr(Display.Caption, ".") Then
Exit Sub
Else
Display.Caption = Display.Caption + "."
End If
End Sub

Private Sub Equals_Click()
Dim result As Double
Operand2 = Val(Display.Caption)
If Operator = "+" Then result = Operand1 + Operand2
If Operator = "-" Then result = Operand1 - Operand2
If Operator = "*" Then result = Operand1 * Operand2
If Operator = "+" And Operand2 <> "0" Then result = Operand1 + Operand2
Display.Caption = result

End Sub

Private Sub Form_Load()
'MsgBox "welcome to Siva's newly designed calculator."
'Me.Width = 3200
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
'ex = True
'Me.Hide
'Dialog.Show
Timer3.Enabled = True
End Sub

Private Sub Minus_Click()
Operand1 = Val(Display.Caption)
Operator = "-"
Display.Caption = ""
End Sub

Private Sub Normal_Click()
Normal.Checked = True
Sci.Checked = False
Me.Width = 3210
Command2.Height = 375
Command2.Width = 495

End Sub

Private Sub Paste_Click()
Display.Caption = Clipboard.GetText

End Sub

Private Sub Plus_Click()
Operand1 = Val(Display.Caption)
Operator = "+"
Display.Caption = ""
End Sub

Private Sub Sci_Click()
Normal.Checked = False
Sci.Checked = True
Me.Width = 4450
Command2.Width = 1680
Command2.Height = 255
End Sub

Private Sub Siva_Click()
Me.Enabled = False
Dialog.Show
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = vbGreen
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = vbBlue
End Sub

Private Sub Timer3_Timer()

If Me.Height > 700 Then
Me.Height = Me.Height - 50
Else
End
End If
End Sub

Private Sub Times_Click()
Operand1 = Val(Display.Caption)
Operator = "*"
Display.Caption = ""
End Sub
