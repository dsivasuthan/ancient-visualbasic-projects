VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Color Mixer - Dhayalan Sivasuthan"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4770
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4680
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   14
      Top             =   3435
      Width           =   4575
   End
   Begin VB.HScrollBar HSGreen 
      Height          =   255
      LargeChange     =   5
      Left            =   720
      Max             =   255
      Min             =   1
      TabIndex        =   2
      Top             =   3120
      Value           =   128
      Width           =   3975
   End
   Begin VB.HScrollBar HSBlue 
      Height          =   255
      LargeChange     =   5
      Left            =   720
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   2760
      Value           =   128
      Width           =   3975
   End
   Begin VB.HScrollBar HSRed 
      Height          =   255
      LargeChange     =   5
      Left            =   720
      Max             =   255
      Min             =   1
      SmallChange     =   5
      TabIndex        =   0
      Top             =   2400
      Value           =   128
      Width           =   3975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   13
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label green 
      BackStyle       =   0  'Transparent
      Caption         =   "128"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label blue 
      BackStyle       =   0  'Transparent
      Caption         =   "128"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label red 
      BackStyle       =   0  'Transparent
      Caption         =   "128"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1890
      Width           =   4455
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Green"
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
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Blue"
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
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
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
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Form_Load()
Label4.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
'ex = False
Label5.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
'Me.Hide
Dialog.Show
'ex = True
Timer2.Enabled = True

End Sub

Private Sub HSBlue_Change()
Label4.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
blue.Caption = HSBlue.Value
Label5.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)

End Sub

Private Sub HSBlue_Scroll()
Label4.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
blue.Caption = HSBlue.Value
Label5.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)

End Sub

Private Sub HSGreen_Change()
Label4.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
green.Caption = HSGreen.Value
Label5.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
End Sub

Private Sub HSGreen_Scroll()
Label4.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
green.Caption = HSGreen.Value
Label5.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)

End Sub

Private Sub HSRed_Change()
Label4.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
red.Caption = HSRed.Value
Label5.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)

End Sub

Private Sub HSRed_Scroll()
Label4.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
red.Caption = HSRed.Value
Label5.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)

End Sub

Private Sub Timer1_Timer()
Label5.ForeColor = vbBlue
End Sub

Private Sub Timer2_Timer()
If Me.Height < 650 Then
End
Else
Me.Height = Me.Height - 10
End If
End Sub
