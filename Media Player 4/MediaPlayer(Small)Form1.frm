VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Media Player"
   ClientHeight    =   1905
   ClientLeft      =   3375
   ClientTop       =   2400
   ClientWidth     =   6090
   Icon            =   "MediaPlayer(Small)Form1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      BackColor       =   &H00808080&
      Caption         =   "^^ Hide Video ^^"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4920
      Width           =   1815
   End
   Begin VB.HScrollBar ScrollVolume 
      Height          =   225
      LargeChange     =   5
      Left            =   4920
      Max             =   100
      TabIndex        =   25
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   3600
      Top             =   2280
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Close"
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
      Left            =   6120
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5640
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   4080
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3720
      Top             =   600
   End
   Begin VB.CommandButton Command10 
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
      Left            =   4800
      TabIndex        =   20
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00808080&
      Caption         =   "Full Screen"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00808080&
      Caption         =   "Show Video"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<"
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
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Mute"
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
      Height          =   195
      Left            =   4920
      TabIndex        =   15
      Top             =   510
      Width           =   735
   End
   Begin VB.ComboBox Rate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "MediaPlayer(Small)Form1.frx":0CCA
      Left            =   4920
      List            =   "MediaPlayer(Small)Form1.frx":0CD7
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   3000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Copies          =   2
      Filter          =   "All media files|*.mp3;*.dat;*.mpg;*.wav;*.avi|MP3|*.mp3|WAV|*.wav|AVI|*.avi|MPEG|*.mpg|DAT|*.dat|"
      FontName        =   "Tahoma"
      FontSize        =   10
      Orientation     =   2
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pause"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Play"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Volume"
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
      Height          =   855
      Left            =   4800
      TabIndex        =   13
      Top             =   0
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Play Rate"
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
      Height          =   615
      Left            =   4800
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   327682
      LargeChange     =   1
      Max             =   300
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      Height          =   330
      Left            =   6120
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2085
      TabIndex        =   21
      Top             =   705
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   3360
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   6000
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   1200
      Top             =   6720
      Width           =   4335
   End
   Begin VB.Label Label6 
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
      Height          =   330
      Left            =   8880
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Elapsed Time"
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
      Left            =   7800
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Height          =   330
      Left            =   960
      TabIndex        =   8
      Top             =   645
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Duration"
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
      Left            =   120
      TabIndex        =   7
      Top             =   675
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   450
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "File"
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
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3225
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5820
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   10266
      _cy             =   5689
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
WindowsMediaPlayer1.settings.mute = True
Else
WindowsMediaPlayer1.settings.mute = False
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
CommonDialog1.ShowOpen
WindowsMediaPlayer1.URL = CommonDialog1.FileName
Label2.Caption = WindowsMediaPlayer1.URL
Label4.Caption = WindowsMediaPlayer1.currentMedia.duration
Timer1.Enabled = True
Slider2.Value = Val(WindowsMediaPlayer1.currentMedia.duration) * 60
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HC000&
End Sub

Private Sub Command10_Click()
Me.Enabled = False
Dialog.Show

End Sub

Private Sub Command11_Click()
'Dim Ex As String
'Ex = MsgBox("Do you really want to quit Siva's Media Player?", vbYesNo, "Siva")
'If Ex = vbYes Then
Timer4.Enabled = True
'End If

End Sub

Private Sub Command12_Click()
Me.Height = 2415
End Sub

Private Sub Command2_Click()
WindowsMediaPlayer1.Controls.play
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HC000&
End Sub

Private Sub Command3_Click()
WindowsMediaPlayer1.Controls.pause

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = &HC000&
End Sub

Private Sub Command4_Click()
WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.BackColor = &HFF&
End Sub

Private Sub Command5_Click()
End
End Sub



Private Sub Command6_Click()
WindowsMediaPlayer1.Controls.fastForward
End Sub

Private Sub Command7_Click()
WindowsMediaPlayer1.Controls.fastForward
End Sub

Private Sub Command8_Click()
If Me.Height = 2415 Then
    Me.Height = 5745
Else
    Me.Height = 2415
End If
End Sub

Private Sub Command9_Click()
On Error Resume Next
WindowsMediaPlayer1.fullScreen = True
End Sub

Private Sub Form_Load()
Me.Height = 2415
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8000000F
Command2.BackColor = &H8000000F
Command3.BackColor = &H8000000F
Command4.BackColor = &H8000000F
'Slider1.Value = WindowsMediaPlayer1.settings.volume
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Command11_Click
Dialog.Show
End Sub

Private Sub Rate_Change()
If Combo2.Text = "" Then
WindowsMediaPlayer1.settings.Rate = "1"
Else
WindowsMediaPlayer1.settings.Rate = Combo2.Text
End If
End Sub

Private Sub Rate_Click()
If Rate.Text = "" Then
WindowsMediaPlayer1.settings.Rate = "1"
Else
WindowsMediaPlayer1.settings.Rate = Rate.Text
End If
End Sub

Private Sub Slider1_Click()
WindowsMediaPlayer1.settings.volume = Slider1.Value
End Sub

Private Sub Slider1_Change()
WindowsMediaPlayer1.settings.volume = Slider2.Value
End Sub



Private Sub Slider1_Scroll()
WindowsMediaPlayer1.settings.volume = Slider2.Value

End Sub

Private Sub ScrollVolume_Change()
WindowsMediaPlayer1.settings.volume = ScrollVolume

End Sub

Private Sub ScrollVolume_Scroll()
WindowsMediaPlayer1.settings.volume = ScrollVolume

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label4.Caption = WindowsMediaPlayer1.currentMedia.durationString

End Sub

Private Sub Timer2_Timer()
Label7.ForeColor = vbBlue
End Sub

Private Sub Timer3_Timer()
Label7.ForeColor = vbGreen
End Sub

Private Sub Timer4_Timer()
If Me.Height > 650 Then
Me.Height = Me.Height - 50
Else
End
End If
End Sub

Private Sub Timer5_Timer()

Slider2.Value = WindowsMediaPlayer1.Controls.currentPosition
Label8.Caption = Slider2.Value
End Sub
