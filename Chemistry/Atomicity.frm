VERSION 5.00
Begin VB.Form frmAtomicity 
   Caption         =   "Atomicity"
   ClientHeight    =   4530
   ClientLeft      =   6975
   ClientTop       =   4290
   ClientWidth     =   3000
   Icon            =   "Atomicity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox NitrogenAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Results"
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
      TabIndex        =   34
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox SulphurAnswer3 
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
      Height          =   285
      Left            =   2480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox PhosphorusAnswer3 
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
      Height          =   285
      Left            =   2480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox OzoneAnswer3 
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
      Height          =   285
      Left            =   2480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox BromineAnswer3 
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
      Height          =   285
      Left            =   2480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox NitrogenAnswer3 
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
      Height          =   285
      Left            =   2480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox ChlorineAnswer3 
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
      Height          =   285
      Left            =   2480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox OxygenAnswer3 
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
      Height          =   285
      Left            =   2475
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox SodiumAnswer3 
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
      Height          =   285
      Left            =   2480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox SulphurAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox PhosphorusAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox OzoneAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox BromineAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox ChlorineAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox OxygenAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox SodiumAnswer2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox SulphurAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox PhosphorusAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   15
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox OzoneAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   14
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox BromineAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox NitrogenAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox ChlorineAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox OxygenAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox SodiumAnswer 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "C/W"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   33
      Top             =   600
      Width           =   375
   End
   Begin VB.Line Line3 
      DrawMode        =   1  'Blackness
      Index           =   2
      X1              =   2280
      X2              =   2280
      Y1              =   480
      Y2              =   3960
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   3000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   1680
      X2              =   1680
      Y1              =   480
      Y2              =   3840
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   3120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Your answer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   32
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Sulphur"
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
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Phosphorus"
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
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Ozone"
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
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Bromine"
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Nitrogen"
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
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Chlorine"
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
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Oxygen"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Sodium"
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
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   6240
      X2              =   10320
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   6240
      X2              =   10320
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   6240
      X2              =   10320
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   6240
      X2              =   10320
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6240
      X2              =   10320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   6240
      X2              =   10320
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6240
      X2              =   10320
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   6240
      X2              =   10320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Type the correct atomicity:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2325
      TabIndex        =   36
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmAtomicity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cl2_Change()

End Sub

Private Sub Command1_Click()
If SodiumAnswer = "1" Then
SodiumAnswer2 = "C"
SodiumAnswer3 = "1"
Else
SodiumAnswer2 = "W"
SodiumAnswer3 = "1"
End If

If OxygenAnswer = "2" Then
OxygenAnswer2 = "C"
OxygenAnswer3 = "2"
Else: OxygenAnswer2 = "W"
OxygenAnswer3 = "2"
End If
If ChlorineAnswer = "2" Then
ChlorineAnswer2 = "C"
ChlorineAnswer3 = "2"
Else
ChlorineAnswer2 = "W"
ChlorineAnswer3 = "2"
End If
If NitrogenAnswer = "2" Then
NitrogenAnswer2 = "C"
NitrogenAnswer3 = "2"
Else: NitrogenAnswer2 = "W"
NitrogenAnswer3 = "2"
End If
If BromineAnswer = "2" Then
BromineAnswer2 = "C"
BromineAnswer3 = "2"
Else: BromineAnswer2 = "W"
BromineAnswer3 = "2"
End If
If OzoneAnswer = "3" Then
OzoneAnswer2 = "C"
OzoneAnswer3 = "3"
Else: OzoneAnswer2 = "W"
OzoneAnswer3 = "3"
End If
If PhosphorusAnswer = "4" Then
PhosphorusAnswer2 = "C"
PhosphorusAnswer3 = "4"
Else: PhosphorusAnswer2 = "W"
PhosphorusAnswer3 = "4"
End If
If SulphurAnswer = "8" Then
SulphurAnswer2 = "C"
SulphurAnswer3 = "8"
Else: SulphurAnswer2 = "W"
SulphurAnswer3 = "8"
End If

End Sub

Private Sub Form_Load()
MsgBox "Good Luck, From Siva", vbOKOnly, "Siva"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
frmWelcomescreen.Show


End Sub
