VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   5115
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   12450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "Form1.frx":0CCA
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5115
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   465
      Left            =   5280
      TabIndex        =   24
      Top             =   4590
      Width           =   6735
   End
   Begin VB.Timer tmrScroll 
      Interval        =   10
      Left            =   5280
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   960
      Top             =   5160
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      LargeChange     =   5
      Left            =   2520
      Max             =   100
      SmallChange     =   5
      TabIndex        =   16
      Top             =   4620
      Value           =   50
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12000
      Top             =   4560
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   165
      LargeChange     =   5
      Left            =   315
      Max             =   200
      MouseIcon       =   "Form1.frx":19C24
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4245
      Width           =   4740
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Playlist"
      ForeColor       =   &H0000FF00&
      Height          =   2055
      Left            =   5280
      TabIndex        =   3
      Top             =   2400
      Width           =   6735
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "Remove Selection"
         Height          =   285
         Left            =   1800
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Clear Playlist"
         Height          =   285
         Left            =   120
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton cmdPlaySel 
         BackColor       =   &H00808080&
         Caption         =   "Play selection"
         Height          =   285
         Left            =   5040
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1695
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Playlist Editor"
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   5280
      TabIndex        =   6
      Top             =   240
      Width           =   6735
      Begin VB.CommandButton cmdAddAllFiles 
         BackColor       =   &H00808080&
         Caption         =   "Add all files"
         Height          =   735
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdAddSelFiles 
         BackColor       =   &H00808080&
         Caption         =   "Add selected file(s)"
         Height          =   660
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1440
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1455
         Left            =   2280
         Pattern         =   "*.mp3;*.wav;*.mpg;*.dat"
         TabIndex        =   8
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox Filter 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   315
         ItemData        =   "Form1.frx":1A4EE
         Left            =   3840
         List            =   "Form1.frx":1A501
         TabIndex        =   7
         Text            =   "All media files"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Filter:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Playlist >>"
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
      Left            =   3960
      MouseIcon       =   "Form1.frx":1A532
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblDuration 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3285
      TabIndex        =   22
      Top             =   3885
      Width           =   1695
   End
   Begin VB.Label lblBar 
      BackStyle       =   0  'Transparent
      Height          =   4335
      Left            =   12120
      MousePointer    =   9  'Size W E
      TabIndex        =   21
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12150
      TabIndex        =   20
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgNextNormal 
      Height          =   330
      Left            =   960
      Picture         =   "Form1.frx":1ADFC
      Top             =   5880
      Width           =   360
   End
   Begin VB.Image imgPreNormal 
      Height          =   375
      Left            =   2400
      Picture         =   "Form1.frx":1B166
      Top             =   6720
      Width           =   360
   End
   Begin VB.Image imgPlaynormal 
      Height          =   450
      Left            =   1200
      Picture         =   "Form1.frx":1B4DB
      Top             =   6720
      Width           =   465
   End
   Begin VB.Image imgStopNormal 
      Height          =   345
      Left            =   2520
      Picture         =   "Form1.frx":1B8E0
      Top             =   6360
      Width           =   330
   End
   Begin VB.Image imgPauseNormal 
      Height          =   450
      Left            =   240
      Picture         =   "Form1.frx":1BC2F
      Top             =   6240
      Width           =   450
   End
   Begin VB.Image imgPlay 
      Height          =   450
      Left            =   1545
      MouseIcon       =   "Form1.frx":1C73B
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1D005
      Top             =   4515
      Width           =   465
   End
   Begin VB.Image imgNext 
      Height          =   330
      Left            =   2010
      MouseIcon       =   "Form1.frx":1D40A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1DCD4
      Top             =   4590
      Width           =   360
   End
   Begin VB.Image imgPrevious 
      Height          =   375
      Left            =   1190
      MouseIcon       =   "Form1.frx":1E03E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1E908
      Top             =   4560
      Width           =   360
   End
   Begin VB.Image imgStop 
      Height          =   345
      Left            =   810
      MouseIcon       =   "Form1.frx":1EC7D
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1F547
      Top             =   4545
      Width           =   330
   End
   Begin VB.Image imgPause 
      Height          =   450
      Left            =   315
      MouseIcon       =   "Form1.frx":1F896
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":20160
      Top             =   4530
      Width           =   450
   End
   Begin VB.Image imgNextMouseDown 
      Height          =   360
      Left            =   600
      Picture         =   "Form1.frx":20C6C
      Top             =   5880
      Width           =   360
   End
   Begin VB.Image imgPlayMouseDown 
      Height          =   450
      Left            =   720
      Picture         =   "Form1.frx":21370
      Top             =   6720
      Width           =   510
   End
   Begin VB.Image imgPreMouseDown 
      Height          =   360
      Left            =   1680
      Picture         =   "Form1.frx":21764
      Top             =   6720
      Width           =   345
   End
   Begin VB.Label lblMediaName 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   345
      TabIndex        =   15
      Top             =   3885
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   280
      X2              =   5075
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   1725
      MouseIcon       =   "Form1.frx":21E68
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":22732
      ToolTipText     =   "Edit Playlist"
      Top             =   60
      Width           =   300
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3120
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   4620
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
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8149
      _cy             =   5503
   End
   Begin VB.Image Image2 
      Height          =   3615
      Left            =   0
      Picture         =   "Form1.frx":22A80
      Top             =   360
      Width           =   5190
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      Height          =   4215
      Left            =   5160
      TabIndex        =   14
      Top             =   240
      Width           =   135
   End
   Begin VB.Image imgStopMouseOver 
      Height          =   450
      Left            =   2040
      Picture         =   "Form1.frx":5FDD4
      Top             =   6240
      Width           =   405
   End
   Begin VB.Image imgStopMouseDown 
      Height          =   450
      Left            =   1680
      Picture         =   "Form1.frx":6016B
      Top             =   6240
      Width           =   390
   End
   Begin VB.Image imgPauseMouseOver 
      Height          =   450
      Left            =   720
      Picture         =   "Form1.frx":604FE
      Top             =   6240
      Width           =   450
   End
   Begin VB.Image imgPauseMouseDown 
      Height          =   450
      Left            =   1200
      Picture         =   "Form1.frx":6100A
      Top             =   6240
      Width           =   450
   End
   Begin VB.Image imgPlayMouseOver 
      Height          =   450
      Left            =   240
      Picture         =   "Form1.frx":61B16
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   4820
      MouseIcon       =   "Form1.frx":61F2E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":62BF8
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image imgNextMouseOver 
      Height          =   360
      Left            =   1320
      Picture         =   "Form1.frx":62F0C
      ToolTipText     =   "Next"
      Top             =   5880
      Width           =   360
   End
   Begin VB.Image imgPreMouseOver 
      Height          =   360
      Left            =   2040
      Picture         =   "Form1.frx":63610
      ToolTipText     =   "Previous"
      Top             =   6720
      Width           =   345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   75
      Width           =   2295
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   4485
      Picture         =   "Form1.frx":63D14
      Top             =   55
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "Form1.frx":64640
      Top             =   0
      Width           =   4485
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   4485
      Picture         =   "Form1.frx":6BA88
      Top             =   0
      Width           =   690
   End
   Begin VB.Image Image3 
      Height          =   1170
      Left            =   0
      Picture         =   "Form1.frx":6CCD8
      Top             =   3960
      Width           =   5190
   End
   Begin VB.Image imgPlaylistDown 
      Height          =   2175
      Left            =   5160
      Picture         =   "Form1.frx":809FC
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   7305
   End
   Begin VB.Image imgPlaylistTop 
      Height          =   2325
      Left            =   5160
      Picture         =   "Form1.frx":8150D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7305
   End
   Begin VB.Menu App 
      Caption         =   "App"
      Visible         =   0   'False
      Begin VB.Menu Move 
         Caption         =   "Move"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Playing As Boolean
Dim Stoped As Boolean
Dim BorderAvailable As Boolean

Private Sub Command1_Click()
List1.Clear
End Sub

Private Sub cmdAddAllFiles_Click()
MousePointer = 11
Dim CountFile, AddedFileCount
CountFile = File1.ListCount
If CountFile <= 0 Then Exit Sub
AddedFileCount = 0
Do Until CountFile = AddedFileCount
List1.AddItem File1.Path & "\" & File1.List(AddedFileCount)
AddedFileCount = AddedFileCount + 1
Loop
MousePointer = 0
End Sub

Private Sub cmdAddSelFiles_Click()
If File1 = "" Then
MsgBox "Make a selection to add a file"
Exit Sub
End If
List1.AddItem Dir1.Path + "\" + File1
End Sub

Private Sub cmdPlaySel_Click()
If List1.ListCount <= 0 Then Exit Sub
WindowsMediaPlayer1.URL = List1.Text
End Sub

Private Sub Command2_Click()
If List1.ListCount = 0 Then Exit Sub
If List1.Text = "" Then
MsgBox "Select the file to be removed from the playlist"
Exit Sub
End If
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Command3_Click()
Label1_Click
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub






Private Sub File1_DblClick()
If File1 = "" Then
MsgBox "Make a selection to add a file"
Exit Sub
End If
List1.AddItem Dir1.Path + "\" + File1
End Sub



Private Sub Form_Click()
If Me.BorderStyle = 0 And Me.Caption = "" Then
PopupMenu App
Me.Height = 5145

Else
If Me.Caption <> "" Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

Else
If Me.BorderStyle = 1 Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

End If
End If
End If

End Sub



Private Sub Form_Load()
Me.Height = 5145
Me.Width = 5190
Drive1.Drive = "e:"
Dir1.Path = "E:\BackUp - 21.12.2007\CD Copies\Siva (F)\Songs"
BorderAvailable = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = imgNextNormal.Picture
imgPlay.Picture = imgPlaynormal.Picture
imgPause.Picture = imgPauseNormal.Picture
imgPrevious.Picture = imgPreNormal.Picture
imgStop.Picture = imgStopNormal.Picture

End Sub


Private Sub HScroll1_Change()
'tmrScroll.Enabled = False
'WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value
'tmrScroll.Enabled = True

End Sub

Private Sub HScroll1_Scroll()
tmrScroll.Enabled = False
WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value
tmrScroll.Enabled = True

End Sub

Private Sub HScroll2_Change()
WindowsMediaPlayer1.settings.volume = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
WindowsMediaPlayer1.settings.volume = HScroll2.Value
End Sub

Private Sub Image1_Click()
If Me.BorderStyle = 0 And Me.Caption = "" Then
PopupMenu App
Me.Height = 5145

Else
If Me.Caption <> "" Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

Else
If Me.BorderStyle = 1 Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

End If
End If
End If
End Sub

Private Sub Image2_Click()
If Me.BorderStyle = 0 And Me.Caption = "" Then
PopupMenu App
Me.Height = 5145

Else
If Me.Caption <> "" Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

Else
If Me.BorderStyle = 1 Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

End If
End If
End If
End Sub

Private Sub Image3_Click()
If Me.BorderStyle = 0 And Me.Caption = "" Then
PopupMenu App
Me.Height = 5145

Else
If Me.Caption <> "" Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

Else
If Me.BorderStyle = 1 Then
Me.BorderStyle = 0
Me.Caption = ""
Me.Height = 5145

End If
End If
End If

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = imgNextNormal.Picture
imgPlay.Picture = imgPlaynormal.Picture
imgPause.Picture = imgPauseNormal.Picture
imgPrevious.Picture = imgPreNormal.Picture
imgStop.Picture = imgStopNormal.Picture
End Sub

Private Sub Image5_Click()
Me.WindowState = 1
End Sub

Private Sub Image6_Click()
End
End Sub

Private Sub Image7_Click()
If Me.Width = 12495 Then
Me.Width = 5190
Else
Me.Width = 12495
End If
'Timer1.Enabled = True
End Sub

Private Sub imgNext_Click()
Me.MousePointer = 11
If List1.ListCount <= 0 Then
MsgBox "There are no media in the playlist to navigate. Click Edit Playlist button on top to edit playlist"
Me.MousePointer = 0
Exit Sub
End If
If List1.ListIndex + 1 = List1.ListCount Then
List1.ListIndex = 0
List1_DblClick
Me.MousePointer = 0
Exit Sub
End If
List1.ListIndex = List1.ListIndex + 1
List1_DblClick
Me.MousePointer = 0
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = imgNextMouseDown.Picture
End Sub

Private Sub imgNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = imgNextMouseOver.Picture

End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = imgNextMouseOver.Picture

End Sub

Private Sub imgNextMouseOver_Click()
If List1.ListCount <= 0 Then Exit Sub
WindowsMediaPlayer1.URL = List1.List(List1.ListIndex)

End Sub




Private Sub imgPause_Click()
WindowsMediaPlayer1.Controls.pause
End Sub



Private Sub imgPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPause.Picture = imgPauseMouseDown
End Sub

Private Sub imgPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPause.Picture = imgPauseMouseOver
End Sub

Private Sub imgPlay_Click()
WindowsMediaPlayer1.Controls.play
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPlay.Picture = imgPlayMouseDown.Picture
End Sub

Private Sub imgPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPlay.Picture = imgPlayMouseOver.Picture

End Sub

Private Sub imgPrevious_Click()
If List1.ListCount <= 0 Then
MsgBox "There are no media in the playlist to navigate. Click Edit Playlist button on top to edit playlist"
Me.MousePointer = 0
Exit Sub
End If
If List1.ListIndex + 1 = List1.ListCount Then
List1.ListIndex = 0
List1_DblClick
Me.MousePointer = 0
Exit Sub
End If
List1.ListIndex = List1.ListIndex - 1
List1_DblClick
Me.MousePointer = 0
End Sub

Private Sub imgPrevious_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPrevious.Picture = imgPreMouseDown.Picture
End Sub

Private Sub imgPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPrevious.Picture = imgPreMouseOver.Picture

End Sub

Private Sub imgPrevious_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPrevious.Picture = imgPreMouseOver.Picture

End Sub

Private Sub imgStop_Click()
WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStop.Picture = imgStopMouseDown.Picture
End Sub

Private Sub imgStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStop.Picture = imgStopMouseOver.Picture
End Sub





Private Sub Label1_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Label5_Click()
If Me.Width = 12495 Then
Label5.Caption = "Playlist >"
Me.Width = 5190
Else
Me.Width = 12495
Label5.Caption = "Playlist <"
End If

End Sub

Private Sub lblBar_Click()
Me.Height = 5145
Me.Width = 5190
End Sub

Private Sub List1_DblClick()
WindowsMediaPlayer1.URL = List1.Text
End Sub

Private Sub Move_Click()
If BorderAvailable = True Then
Me.BorderStyle = 0
Me.Caption = ""
Else
If BorderAvailable = False Then
Me.BorderStyle = 1
Me.Caption = "Dhayalan Sivasuthan - My Media Player"
End If
End If
End Sub

Private Sub Timer1_Timer()
If Me.Width > 12470 Then
Me.Width = 12470
Timer1.Enabled = False
Else
Me.Width = Me.Width + 10
End If
End Sub

Private Sub Timer2_Timer()
If WindowsMediaPlayer1.URL = "" Then Exit Sub
lblDuration.Caption = WindowsMediaPlayer1.Controls.currentPositionString & "/" & WindowsMediaPlayer1.currentMedia.durationString
If WindowsMediaPlayer1.currentMedia.duration < 20 Then Exit Sub
If WindowsMediaPlayer1.Controls.currentPosition > WindowsMediaPlayer1.currentMedia.duration - 20 Then
imgNext_Click
End If

End Sub

Private Sub tmrScroll_Timer()
HScroll1.Value = WindowsMediaPlayer1.Controls.currentPosition
End Sub

Private Sub WindowsMediaPlayer1_MediaChange(ByVal Item As Object)
lblMediaName.Caption = WindowsMediaPlayer1.currentMedia.Name
HScroll1.Max = WindowsMediaPlayer1.currentMedia.duration
End Sub



Private Sub WindowsMediaPlayer1_PlayStateChange(ByVal NewState As Long)
Text1.Text = NewState 'WindowsMediaPlayer1.Controls.currentPositionString

End Sub
