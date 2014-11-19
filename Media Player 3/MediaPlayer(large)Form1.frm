VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Media Player 3"
   ClientHeight    =   5610
   ClientLeft      =   1215
   ClientTop       =   630
   ClientWidth     =   5415
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MediaPlayer(large)Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<  Hide"
      Height          =   375
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5130
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show Video"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Default"
      Height          =   300
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4590
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Time"
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   1920
      TabIndex        =   28
      Top             =   3960
      Width           =   1575
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         X1              =   120
         X2              =   1440
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Went time"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Play"
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
      Left            =   8400
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mp3"
      DialogTitle     =   "Open"
      Filter          =   "*.mp3"
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Open Dialog"
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
      Left            =   8520
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Play"
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
      Left            =   7560
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Open avi file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8280
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Settings"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton Command5 
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
      Left            =   7080
      TabIndex        =   18
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00808080&
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   6600
      Top             =   2520
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Play Rate"
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   3600
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
      Begin VB.ComboBox PlayRate 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   315
         ItemData        =   "MediaPlayer(large)Form1.frx":0CCA
         Left            =   120
         List            =   "MediaPlayer(large)Form1.frx":0CD7
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   2520
   End
   Begin VB.Timer Timer4 
      Interval        =   2000
      Left            =   3840
      Top             =   5880
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   4320
      Top             =   5880
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6480
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   1440
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00808080&
      Caption         =   "Stop"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
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
      Left            =   9480
      TabIndex        =   5
      Top             =   6720
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   315
      ItemData        =   "MediaPlayer(large)Form1.frx":0CE4
      Left            =   3360
      List            =   "MediaPlayer(large)Form1.frx":0CF7
      TabIndex        =   4
      Text            =   "*.mp3;*.wav;*.mpg"
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox filFile 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   1845
      Left            =   2760
      Pattern         =   "*.mp3;*.wav;*.mpg"
      ReadOnly        =   0   'False
      System          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.DirListBox DirDirectories 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox drvDrive 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Volume"
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Mute"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   1395
      End
      Begin VB.HScrollBar Slider1 
         Height          =   255
         LargeChange     =   5
         Left            =   240
         Max             =   50
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   3375
      Left            =   7440
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5953
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   449
      FullHeight      =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   6600
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "File:"
      Height          =   255
      Left            =   6120
      TabIndex        =   21
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   6480
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Name"
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   6240
      Top             =   4800
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label Label7 
      Caption         =   "Duration:"
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Time:"
      Height          =   255
      Left            =   7320
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Dhayalan Sivasuthan"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   450
      Left            =   120
      TabIndex        =   7
      Top             =   2970
      Width           =   5175
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3720
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   5220
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
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9208
      _cy             =   6562
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

Private Sub cmdRefresh_Click()
filFile.Pattern = Combo1.Text
End Sub

Private Sub cmdStop_Click()
WindowsMediaPlayer1.Controls.Stop
Command6.Caption = "Play"
End Sub

Private Sub cmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdStop.BackColor = vbBlue
End Sub

Private Sub Combo1_Change()
filFile.Pattern = Combo1.Text
End Sub

Private Sub Combo1_Click()
filFile.Pattern = Combo1.Text
End Sub


Private Sub Command1_Click()
On Error Resume Next
'If Animation1.Visible = True Then
'Animation1.Open lblAddress.Caption
'Else
If filFile = "" Then
MsgBox "Error... No selection made"
Else
WindowsMediaPlayer1.URL = lblAddress.Caption
Timer5.Enabled = True
Label8.Caption = WindowsMediaPlayer1.currentMedia.Name
Command6.Caption = "Pause"
End If
'End If
End Sub '

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.BackColor = vbGreen
End Sub

Private Sub Command10_Click()
CommonDialog1.ShowOpen
lblAddress.Caption = CommonDialog1.FileName
End Sub

Private Sub Command11_Click()
WindowsMediaPlayer1.URL = lblAddress
CommonDialog1.FileName = ""
End Sub

Private Sub Command12_Click()
'If
Me.Width = 10815  'Then
'Me.Height = 5460

End Sub

Private Sub Command13_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Command15_Click()
Me.Width = 5460
End Sub

Private Sub Command2_Click()
'Dim Ex As Integer
'Ex = MsgBox("Do you really want to exit from the media player?", vbYesNo, "Siva's Media Player")
'If Ex = vbYes The

Timer1.Enabled = True
'End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = vbRed
End Sub


Private Sub Command4_Click()
CommonDialog1.ShowOpen
End Sub

Private Sub Command6_Click()
If Command6.Caption = "Play" Then
WindowsMediaPlayer1.Controls.Play
Command6.Caption = "Pause"
ElseIf Command6.Caption = "Pause" Then
WindowsMediaPlayer1.Controls.pause
Command6.Caption = "Play"
End If

If WindowsMediaPlayer1.URL = "" Then
Command6.Caption = "Play"
End If
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.BackColor = vbBlue
End Sub

Private Sub Command7_Click()
If Form1.Height = 7890 Then
Form1.Height = 6700
Else
Form1.Height = 7890
End If
End Sub

Private Sub Command8_Click()
On Error Resume Next
'Picture1.Picture = lblAddress
'Animation1.Open lblAddress.Caption
Animation1.Visible = True
Combo1.Text = "*.avi"
End Sub

Private Sub Command9_Click()
On Error Resume Next
Animation1.Play
End Sub

Private Sub DirDirectories_Change()
filFile.Path = DirDirectories.List(DirDirectories.ListIndex)

End Sub

Private Sub drvDrive_Change()
DirDirectories.Path = drvDrive.Drive
End Sub

Private Sub filFile_Click()
lblAddress.Caption = DirDirectories.Path + "\" + filFile
End Sub



Private Sub Form_Load()
'DirDirectories.Path = "e:\My Documents\My Music\Songs\Mix"
Me.Width = 5460
Slider1.Value = WindowsMediaPlayer1.settings.volume

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = vbButtonFace
cmdStop.BackColor = vbButtonFace
Command6.BackColor = vbButtonFace
Command1.BackColor = vbButtonFace
End Sub

'Private Sub Form_Terminate()
'Dim Ex As Integer
'Ex = MsgBox("Do you really want to exit from the media player?", vbYesNo, "Siva's Media Player")
'If Ex = vbYes Then
'Timer1.Enabled = True
'Else
'Form1.Hide
'Dialog.Show
'End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
'Dim Ex As Integer
'Ex = MsgBox("Do you really want to exit from the media player?", vbYesNo, "Siva's Media Player")
'If Ex = vbYes Then
Command2_Click
'End If
End Sub

Private Sub Slider1_Click()

End Sub

Private Sub PlayRate_Change()
If PlayRate.Text = "" Then
WindowsMediaPlayer1.settings.Rate = "1"
Else
WindowsMediaPlayer1.settings.Rate = PlayRate.Text
End If
End Sub

Private Sub PlayRate_Click()
If PlayRate.Text = "" Then
WindowsMediaPlayer1.settings.Rate = "1"
Else
WindowsMediaPlayer1.settings.Rate = PlayRate.Text
End If

End Sub

Private Sub Slider1_Change()
WindowsMediaPlayer1.settings.volume = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
WindowsMediaPlayer1.settings.volume = Slider1.Value
End Sub

Private Sub Timer1_Timer()
If Form1.Height > 620 Then
Me.WindowState = Normal
Form1.Height = Form1.Height - 68
Else
End
'Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If Form1.Width > 1770 Then
Form1.Width = Form1.Width - 50
Else
End
End If
End Sub

Private Sub Timer3_Timer()
Label3.ForeColor = vbGreen
End Sub

Private Sub Timer4_Timer()
Label3.ForeColor = vbBlue
End Sub

Private Sub Timer5_Timer()
Label4.Caption = WindowsMediaPlayer1.currentMedia.duration
End Sub

Private Sub Timer6_Timer()
Label5.Caption = WindowsMediaPlayer1.Controls.currentPosition
End Sub
