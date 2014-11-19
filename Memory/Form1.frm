VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Monitor - Dhayalan Sivasuthan"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mini mode"
      Height          =   375
      Left            =   3120
      TabIndex        =   28
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Frame Frame4 
      Caption         =   "Memory Usage"
      Height          =   5415
      Left            =   4680
      TabIndex        =   24
      Top             =   1080
      Width           =   1335
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   4935
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   8705
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Virtual Memory"
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   4455
      Begin VB.Label Label22 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Used Virtual Memory"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Total Virtual Memory"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "Available Virtual Memory"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Page File"
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   4455
      Begin VB.Label Label4 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Total Page file"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Available Page file"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label20 
         Caption         =   "Used Page file"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label21 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Physical Memory"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
      Begin VB.Label Label2 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Total Physical Memory"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Available Physical Memory"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "gfgfgfgfgfgfgfgfgfg"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Used Physical Memory"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   6720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Full mode"
      Height          =   255
      Left            =   4440
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label17 
      Caption         =   "Memory Monitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label16 
      Caption         =   "Dhayalan Sivasuthan's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "Form1.frx":0ECA
      Top             =   45
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "gfgfgfgfgfgfgfgfgfg"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Memory Load"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dialog.Show
Me.Hide
Me.Enabled = False
End Sub


Private Sub Command2_Click()
Me.Hide
Form2.Show
'Me.Height = 1830
'Frame1.Visible = False
'Frame2.Visible = False
'Frame3.Visible = False
'Frame4.Visible = False
'ProgressBar2.Visible = True
'Command4.Visible = True
End Sub

Private Sub Command3_Click()
Command1_Click
End Sub

Private Sub Command4_Click()
Me.Height = 7440
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Frame4.Visible = True
ProgressBar2.Visible = False
Command4.Visible = False
End Sub

Private Sub Form_Load()

StayOnTop Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dialog.Show
Me.Hide
DEx = True
End Sub

Private Sub Timer1_Timer()
Dim memStat As MEMORYSTATUS
memStat.dwlength = Len(memStat)
Call GlobalMemoryStatus(memStat)
Label1.Caption = memStat.dwMemoryLoad
Label2.Caption = memStat.dwTotalPhys / 1024 / 1024 & " MB"
Label3.Caption = memStat.dwAvailPhys / 1024 / 1024 & " MB"
Label4.Caption = memStat.dwTotalPageFile / 1024 / 1024 & " MB"
Label5.Caption = memStat.dwAvailPage / 1024 / 1024 & " MB"
Label6.Caption = memStat.dwTotalVirtual / 1024 / 1024 & " MB"
Label7.Caption = memStat.dwAvailVirtual / 1024 / 1024 & " MB"
ProgressBar1.Max = Val(Label2.Caption)
ProgressBar1.Value = Val(Label2.Caption) - Val(Label3.Caption)
Label18.Caption = (Val(Label2.Caption) - Val(Label3.Caption)) & " MB"
Label21.Caption = (Val(Label4.Caption) - Val(Label5.Caption)) & " MB"
Label22.Caption = (Val(Label6.Caption) - Val(Label7.Caption)) & " MB"
Form2.ProgressBar1.Max = Val(Label2.Caption)
'Form2.ProgressBar2.Max = Val(Label2.Caption)
Form2.ProgressBar1.Value = Val(Label2.Caption) - Val(Label3.Caption)
'Form2.ProgressBar2.Value = Val(Label21.Caption)
Form2.Label3.Caption = Format((Form2.ProgressBar1.Value / Form2.ProgressBar1.Max), "0.00") * 100 & " %"
End Sub
