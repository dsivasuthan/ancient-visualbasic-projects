VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer"
   ClientHeight    =   1680
   ClientLeft      =   8610
   ClientTop       =   6630
   ClientWidth     =   1680
   Icon            =   "TimerForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   1200
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3360
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   1440
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   60
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      Height          =   255
      Left            =   120
      Top             =   720
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   300
      Left            =   120
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label d 
      Alignment       =   2  'Center
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label c 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label b 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Result, UsedTime, loca
UsedTime = ((Val(c.Caption) * 60) * 60) + (Val(b.Caption) * 60) + Val(a.Caption)
loca = App.Path & "\" & "TMR.txt"
Open loca For Input As #1
Line Input #1, Result
Close #1
loca = App.Path & "\" & "TMR.txt"
Open loca For Output As #1
Result = Val(Result) + Val(UsedTime)
Print #1, Result
Close #1


MsgBox UsedTime & " seconds"

a = ""
b = ""
c = ""
Command3.Caption = "Pause"
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
'Dim Ex As Integer
'Ex = MsgBox("Do you wanted to exit?", vbYesNo, "Siva's Timer")
'If Ex = vbYes Then
Dim Result, UsedTime, loca
UsedTime = ((Val(c.Caption) * 60) * 60) + (Val(b.Caption) * 60) + Val(a.Caption)
loca = App.Path & "\" & "TMR.txt"
Open loca For Input As #1
Line Input #1, Result
Close #1
loca = App.Path & "\" & "TMR.txt"
Open loca For Output As #1
Result = Val(Result) + Val(UsedTime)
Print #1, Result
Close #1

Timer4.Enabled = True
MsgBox UsedTime & " seconds"

'End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Pause" Then
Command3.Caption = "Continue"
Timer1.Enabled = False
Else
If Command3.Caption = "Continue" Then
Command3.Caption = "Pause"
Timer1.Enabled = True
End If
End If
End Sub

Private Sub Command4_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Form_Load()
MsgBox "Welcome! Place me in one of the corners. I will show how long you have used you computer since you last switched on the PC. OK?", vbOKOnly, "Siva's timer"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
'MsgBox App.Path

Command2_Click
End Sub

Private Sub Timer1_Timer()
a = Val(a) + 1
If Val(a) = 59 Then
a = "0"
b = Val(b) + 1
End If
If Val(b) = 59 Then
b = "0"
c = Val(c) + 1
End If
Label1 = Time
End Sub

Private Sub Timer2_Timer()
Label2 = "Dhayalan Sivasuthan"

End Sub

Private Sub Timer3_Timer()
Label2 = ""
Timer2.Enabled = True
End Sub

Private Sub Timer4_Timer()
If Me.Height > 550 Then
Me.Height = Me.Height - 5
Else
End
End If
End Sub
