VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{91AB9C8C-C44B-11D2-ACDB-444553540000}#1.0#0"; "ScreenCapture.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Capture"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
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
   ScaleHeight     =   2985
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2280
      Top             =   1200
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Include me"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin ARButtonCtrl.ARButton ARButton4 
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Capture And Save"
      ForeColor       =   16711680
      BackColorOnMouse=   -2147483633
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   3
      Picture         =   "Form1.frx":0CCA
   End
   Begin ARButtonCtrl.ARButton ARButton3 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "About"
      ForeColor       =   -2147483630
      BackColorOnMouse=   -2147483633
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show preview"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin ARButtonCtrl.ARButton ARButton2 
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   1560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      Caption         =   "Change"
      ForeColor       =   -2147483630
      BackColorOnMouse=   -2147483633
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "C:\Capture.bmp"
      Top             =   1560
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp"
      DialogTitle     =   "Save As"
      FileName        =   "Capture"
      Filter          =   "Bitmap File|*.bmp"
   End
   Begin ScrCapture.ScreenCapture ScreenCapture1 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label3 
      Caption         =   "Save As:"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Screen Capture "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   1665
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Dhayalan Sivasuthan's"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   180
      Width           =   3855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":19A4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ARButton2_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
Text1.Text = CommonDialog1.FileName
Else
Text1.Text = "C:\Capture.bmp"
End If
End Sub

Private Sub ARButton3_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub ARButton4_Click()
If Check2.Value = 0 Then
Me.WindowState = 1
ScreenCapture1.FileName = Text1.Text
ScreenCapture1.StartCapture = True
Else
Me.WindowState = 0
ScreenCapture1.FileName = Text1.Text
If Check1.Value = 0 Then
ScreenCapture1.Preview = False
Else
ScreenCapture1.Preview = True
End If
ScreenCapture1.StartCapture = True
End If
Me.WindowState = 0
If Check1.Value = 0 Then
ScreenCapture1.Preview = False
Else
ScreenCapture1.Preview = True
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
ScreenCapture1.Preview = False
Else
ScreenCapture1.Preview = False
End If
End Sub

Private Sub Form_Load()
'Set Skinner1.Forms = Forms
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = RGB(Rnd * 500, Rnd * 500, Rnd * 500)
Label2.ForeColor = RGB(Rnd * 500, Rnd * 500, Rnd * 500)
End Sub
