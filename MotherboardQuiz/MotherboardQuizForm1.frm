VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hardware - Motherboard Naming"
   ClientHeight    =   8910
   ClientLeft      =   2415
   ClientTop       =   2265
   ClientWidth     =   8310
   Icon            =   "MotherboardQuizForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "MotherboardQuizForm1.frx":08CA
   ScaleHeight     =   8910
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   4320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
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
      Left            =   6840
      TabIndex        =   36
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   6840
      TabIndex        =   35
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox P 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   33
      Top             =   8520
      Width           =   495
   End
   Begin VB.TextBox I 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   8
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox N 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   29
      Top             =   8040
      Width           =   495
   End
   Begin VB.TextBox C 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   2
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox L 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   11
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox K 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   10
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox G 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   6
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox H 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   7
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox O 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   32
      Top             =   8520
      Width           =   495
   End
   Begin VB.TextBox D 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   3
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox A 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   0
      Top             =   8040
      Width           =   495
   End
   Begin VB.TextBox E 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   4
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox M 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   28
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox B 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   1
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox J 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   9
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox F 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   5
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   34
      Top             =   6720
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   120
      Picture         =   "MotherboardQuizForm1.frx":0C0C
      ScaleHeight     =   6075
      ScaleWidth      =   8070
      TabIndex        =   12
      Top             =   480
      Width           =   8130
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   4080
      Y1              =   6840
      Y2              =   8880
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Northbridge Chipset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   31
      Top             =   8520
      Width           =   1860
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Southbridge Chipset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   30
      Top             =   8520
      Width           =   1905
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "CNR Port"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   27
      Top             =   8040
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "CMOS Battery"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   26
      Top             =   7560
      Width           =   1290
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "IDE Cable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   25
      Top             =   7320
      Width           =   870
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Floppy Disk Connector"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   24
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "RAM Slots"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   23
      Top             =   8280
      Width           =   960
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "CPU Socket"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   22
      Top             =   6840
      Width           =   1050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "PCI Slots"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   21
      Top             =   7080
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "AGP Slot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   20
      Top             =   6840
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ATX Power Connector"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   19
      Top             =   7800
      Width           =   2130
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "COM Ports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   18
      Top             =   8280
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "IDE Cable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   17
      Top             =   7560
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "USB (Universal Serial Bus) Ports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   16
      Top             =   7320
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Printer Ports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   15
      Top             =   7800
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mouse and Keyboard Ports (PS/2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   14
      Top             =   8040
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "Type the capital letter in the motherboard picture in textboxes in front of the correct labels"
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
      TabIndex        =   13
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If A = "A" Then
A = "C"
Else
A = "w"
End If

If B = "B" Then
B = "C"
Else
B = "w"
End If


If C = "C" Then
C = "C"
Else
C = "w"
End If

If D = "D" Then
D = "C"
Else
D = "w"
End If


If E = "E" Then
E = "C"
Else
E = "w"
End If

If F = "F" Then
F = "C"
Else
F = "w"
End If


If G = "G" Then
G = "C"
Else
G = "w"
End If

If H = "H" Then
H = "C"
Else
H = "w"
End If

If I = "I" Then
I = "C"
Else
I = "w"
End If

If J = "J" Then
J = "C"
Else
J = "w"
End If

If K = "K" Then
K = "C"
Else
K = "w"
End If

If L = "L" Then
L = "C"
Else
L = "w"
End If

If M = "M" Then
M = "C"
Else
M = "w"
End If

If N = "N" Then
N = "C"
Else
N = "w"
End If

If O = "O" Then
O = "C"
Else
O = "w"
End If

If P = "P" Then
P = "C"
Else
P = "w"
End If


End Sub

Private Sub Command2_Click()
A = ""
B = ""
C = ""
D = ""
E = ""
F = ""
G = ""
H = ""
I = ""
J = ""
K = ""
L = ""
M = ""
N = ""
O = ""
P = ""
F.SetFocus
End Sub

Private Sub Command3_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Form_Load()
MsgBox "Welcome to Siva's Motherboard Quiz. Good Luck", vbOKOnly, "Siva"
'Form1.F.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
'Dim Ex As String
'Ex = MsgBox("Do you really want to quit Siva's Motherboard Quiz?", vbYesNo, "Siva")
'If Ex = vbYes Then
Timer3.Enabled = True
Dialog.Show
'End If
End Sub

Private Sub Timer3_Timer()
If Me.Height > 700 Then
Me.Height = Me.Height - 60
Form1.WindowState = Normal
Else
End
End If
End Sub
