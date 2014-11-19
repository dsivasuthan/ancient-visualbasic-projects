VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "My Web Browser - Dhayalan Sivasuthan"
   ClientHeight    =   8745
   ClientLeft      =   9600
   ClientTop       =   6435
   ClientWidth     =   9570
   Icon            =   "WebBrowserForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   8400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.PictureBox Slider2 
      Height          =   675
      Left            =   4200
      ScaleHeight     =   615
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      Picture         =   "WebBrowserForm1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   8160
      Top             =   8400
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Search"
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
      Left            =   7560
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7800
      Top             =   8400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      Picture         =   "WebBrowserForm1.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Picture         =   "WebBrowserForm1.frx":1E65
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      Picture         =   "WebBrowserForm1.frx":2340
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Picture         =   "WebBrowserForm1.frx":284C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
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
      Left            =   8850
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1035
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.*"
   End
   Begin VB.TextBox Text1 
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
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   7935
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      OLEDropMode     =   1  'Manual
      Picture         =   "WebBrowserForm1.frx":2CB8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1470
      Width           =   9375
      ExtentX         =   16536
      ExtentY         =   12091
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Save 
         Caption         =   "Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu Navigate 
         Caption         =   "Navigate"
         Begin VB.Menu Back 
            Caption         =   "Back"
         End
         Begin VB.Menu Forward 
            Caption         =   "Forward"
         End
         Begin VB.Menu Home 
            Caption         =   "Home Page"
         End
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
   Begin VB.Menu Time 
      Caption         =   "                        "
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Back_Click()
Command6_Click
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
Command2_Click
End Sub

Private Sub Command10_Click()

'Text1.Text =
WebBrowser1.QueryStatusWB OLECMDID_ENABLE_INTERACTION
End Sub


Private Sub Command2_Click()
WebBrowser1.Navigate Text1.Text
End Sub

Private Sub Command3_Click()
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command4_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub Command5_Click()
On Error Resume Next
WebBrowser1.GoHome
End Sub

Private Sub Command6_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub

Private Sub Command7_Click()
WebBrowser1.Refresh
End Sub

Private Sub Command8_Click()
WebBrowser1.ExecWB OLECMDID_STOP, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command9_Click()
'WebBrowser1.GoSearch
End Sub

Private Sub Form_Load()
WebBrowser1.GoHome


End Sub

Private Sub Form_Resize()
WebBrowser1.Height = Form1.Height - 2500
WebBrowser1.Width = Form1.Width - 350
Command2.Left = Form1.Width - 900
Text1.Width = Form1.Width - 2000
'Line1.X2 = Form1.Width - 250
End Sub

Private Sub Forward_Click()
Command4_Click
End Sub

Private Sub Home_Click()
Command5_Click
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Save_Click()
Command3_Click
End Sub

Private Sub Siva_Click()
About_Click
End Sub

Private Sub Timer1_Timer()
Time.Caption = Now
'Label4.ForeColor = vbBlue
End Sub

Private Sub Timer2_Timer()
'Label4.ForeColor = vbRed
End Sub


Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
Text1.Text = WebBrowser1.LocationURL

End Sub


Private Sub WebBrowser1_DownloadBegin()
On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress
End Sub

Private Sub WebBrowser1_DownloadComplete()
On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress
End Sub
