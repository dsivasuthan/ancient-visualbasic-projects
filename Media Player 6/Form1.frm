VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   9000
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9000
      Top             =   4320
   End
   Begin VB.PictureBox picMax 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5310
      Picture         =   "Form1.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   510
      TabIndex        =   5
      ToolTipText     =   "Fullscreen view"
      Top             =   3045
      Width           =   510
   End
   Begin VB.PictureBox picMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5205
      Picture         =   "Form1.frx":44B0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      ToolTipText     =   "About"
      Top             =   3630
      Width           =   495
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5370
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Form1.frx":48B7
      ScaleHeight     =   480
      ScaleWidth      =   510
      TabIndex        =   2
      ToolTipText     =   "Close"
      Top             =   2430
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   675
      Left            =   6600
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox picMainSkin 
      Height          =   6255
      Left            =   0
      Picture         =   "Form1.frx":8034
      ScaleHeight     =   6195
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.HScrollBar HScroll2 
         Height          =   225
         LargeChange     =   10
         Left            =   4045
         Max             =   100
         SmallChange     =   5
         TabIndex        =   15
         Top             =   4350
         Value           =   50
         Width           =   960
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3830
         TabIndex        =   16
         ToolTipText     =   "Mute/Unmute"
         Top             =   4350
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   1800
         TabIndex        =   13
         Top             =   5760
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   165
         Left            =   1800
         TabIndex        =   12
         Top             =   3960
         Width           =   3165
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000000&
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
         Height          =   315
         ItemData        =   "Form1.frx":73E82
         Left            =   3000
         List            =   "Form1.frx":73E9B
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Media files filter"
         Top             =   435
         Width           =   1860
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
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
         Height          =   1425
         Left            =   825
         TabIndex        =   10
         Top             =   2040
         Width           =   4335
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00000000&
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
         Height          =   1260
         Left            =   3000
         Pattern         =   "*.mp3;*.dat;*.mpg;*.mid;*.wav"
         TabIndex        =   9
         Top             =   765
         Width           =   2100
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00000000&
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
         Height          =   1215
         Left            =   885
         TabIndex        =   8
         Top             =   795
         Width           =   2100
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00000000&
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
         Height          =   315
         Left            =   1125
         TabIndex        =   7
         Top             =   435
         Width           =   1860
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":73EDA
         ToolTipText     =   "About"
         Top             =   2400
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   165
         Picture         =   "Form1.frx":742E1
         ToolTipText     =   "Show Video"
         Top             =   3000
         Width           =   540
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   240
         Picture         =   "Form1.frx":74708
         ToolTipText     =   "Show Playlist"
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   3600
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   4425
         Picture         =   "Form1.frx":74AEF
         Top             =   3570
         Width           =   750
      End
      Begin VB.Image imgNext 
         Height          =   285
         Left            =   3030
         Picture         =   "Form1.frx":74E60
         ToolTipText     =   "Next media"
         Top             =   4320
         Width           =   510
      End
      Begin VB.Image imgPre 
         Height          =   300
         Left            =   2235
         Picture         =   "Form1.frx":75413
         ToolTipText     =   "Previous media"
         Top             =   4320
         Width           =   525
      End
      Begin VB.Image imgStop 
         Height          =   690
         Left            =   1365
         Picture         =   "Form1.frx":759BF
         ToolTipText     =   "Stop"
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image imgPlay 
         Appearance      =   0  'Flat
         Height          =   900
         Left            =   840
         Picture         =   "Form1.frx":760FA
         ToolTipText     =   "Play/Pause"
         Top             =   4200
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "No Media Loaded"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   3600
         Width           =   2655
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   3000
         Left            =   1035
         TabIndex        =   3
         Top             =   435
         Width           =   3780
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
         _cx             =   6668
         _cy             =   5292
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   7560
      TabIndex        =   6
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Style           =   2
      TabFixedWidth   =   2999
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Playlist"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Video"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgPreDown 
      Height          =   285
      Left            =   1800
      Picture         =   "Form1.frx":76C98
      Top             =   8160
      Width           =   510
   End
   Begin VB.Image imgNextDown 
      Height          =   285
      Left            =   2400
      Picture         =   "Form1.frx":77250
      Top             =   8160
      Width           =   510
   End
   Begin VB.Image imgNextUp 
      Height          =   285
      Left            =   2400
      Picture         =   "Form1.frx":77803
      Top             =   7800
      Width           =   510
   End
   Begin VB.Image imgPreUp 
      Height          =   300
      Left            =   1800
      Picture         =   "Form1.frx":77DB6
      Top             =   7800
      Width           =   525
   End
   Begin VB.Image imgStopUp 
      Height          =   690
      Left            =   5040
      Picture         =   "Form1.frx":78362
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgStopDown 
      Height          =   690
      Left            =   4560
      Picture         =   "Form1.frx":78A9D
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image imgPlayMouseOver 
      Height          =   900
      Left            =   1320
      Picture         =   "Form1.frx":791E1
      Top             =   6360
      Width           =   750
   End
   Begin VB.Image imgPlayLost 
      Height          =   900
      Left            =   360
      Picture         =   "Form1.frx":79D72
      Top             =   6360
      Width           =   750
   End
   Begin VB.Image imgPlayMouseDown 
      Height          =   900
      Left            =   2280
      Picture         =   "Form1.frx":7A910
      Top             =   6360
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Playing As Boolean

Private Sub Check1_Click()
If Check1.Value = Checked Then
WindowsMediaPlayer1.settings.mute = False
Else
WindowsMediaPlayer1.settings.mute = True
End If
End Sub

Private Sub Combo1_Change()
If Combo1.Text = "All Media Files" Then
File1.Pattern = "*.mp3;*.dat;*.mpg;*.mid;*.wav"
Else
File1.Pattern = Combo1.Text
End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub



Private Sub File1_DblClick()
List1.AddItem File1.Path + "\" + File1.FileName
End Sub

Private Sub Form_Load()
'on error resume next
    Dim WindowRegion As Long
    
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Set picMainSkin.Picture = LoadPicture(App.Path + "\Background.bmp")
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
Combo1.ListIndex = 0
'Drive1.Drive = "e:"
'Dir1.Path = "E:\BackUp\CD Copies\Siva (F)\Songs"

End Sub






Private Sub HScroll1_Scroll()
Timer1.Enabled = False
WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value
Timer1.Enabled = True

End Sub

Private Sub HScroll2_Change()
WindowsMediaPlayer1.settings.volume = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
WindowsMediaPlayer1.settings.volume = HScroll2.Value
End Sub







Private Sub Image2_Click()
WindowsMediaPlayer1.Visible = False
Dir1.Visible = True
File1.Visible = True
Drive1.Visible = True
Combo1.Visible = True
List1.Visible = True
End Sub

Private Sub Image3_Click()
WindowsMediaPlayer1.Visible = True
Dir1.Visible = False
File1.Visible = False
Drive1.Visible = False
Combo1.Visible = False
List1.Visible = False

End Sub

Private Sub Image4_Click()
Dialog.Show
Me.Enabled = False
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
imgNext.Picture = imgNextDown.Picture
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = imgNextUp.Picture
End Sub

Private Sub imgPlay_Click()
If Playing = True Then
WindowsMediaPlayer1.Controls.pause
Playing = False
Else
WindowsMediaPlayer1.Controls.play
Playing = True
End If

End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPlay.Picture = imgPlayMouseDown.Picture
End Sub




Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPlay.Picture = imgPlayMouseOver.Picture
End Sub

Private Sub imgPre_Click()
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

Private Sub imgPre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPre.Picture = imgPreDown.Picture
End Sub

Private Sub imgPre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPre.Picture = imgPreUp.Picture
End Sub

Private Sub imgStop_Click()
WindowsMediaPlayer1.Controls.stop

End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStop.Picture = imgStopDown.Picture
End Sub


Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStop.Picture = imgStopUp.Picture
End Sub

Private Sub List1_Click()
List1.ToolTipText = List1.Text
End Sub

Private Sub List1_DblClick()
WindowsMediaPlayer1.URL = List1.Text
End Sub



Private Sub List1_GotFocus()
List1.ToolTipText = List1.Text

End Sub

Private Sub picClose_Click()
DEx = True
Me.Hide
Dialog.Show
End Sub

Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub




Private Sub picMainSkin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'imgStop.Height = imgStop.Height + 20
'imgStop.Width = imgStop.Width + 20
'imgPlay.Picture = imgPlayLost.Picture
End Sub

Private Sub picMax_Click()
WindowsMediaPlayer1.fullScreen = True
End Sub

Private Sub picMin_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Timer1_Timer()
HScroll1.Value = WindowsMediaPlayer1.Controls.currentPosition
End Sub

Private Sub Timer2_Timer()
If WindowsMediaPlayer1.URL = "" Then Exit Sub
Label2.Caption = WindowsMediaPlayer1.Controls.currentPositionString
'& "/" &
Label3.Caption = WindowsMediaPlayer1.currentMedia.durationString
If WindowsMediaPlayer1.currentMedia.duration < 5 Then Exit Sub
If WindowsMediaPlayer1.Controls.currentPosition > WindowsMediaPlayer1.currentMedia.duration - 5 Then
imgNext_Click
End If

End Sub

Private Sub WindowsMediaPlayer1_MediaChange(ByVal Item As Object)
Label1.Caption = WindowsMediaPlayer1.currentMedia.Name
HScroll1.Max = WindowsMediaPlayer1.currentMedia.duration
End Sub

