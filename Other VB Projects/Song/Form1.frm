VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
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
   ScaleHeight     =   7965
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6480
      Top             =   5040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5520
      Top             =   5280
   End
   Begin VB.PictureBox picMainSkin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   0
      Picture         =   "Form1.frx":57E2
      ScaleHeight     =   7035
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin Project1.XPButton XPButton8 
         Height          =   300
         Left            =   5400
         TabIndex        =   15
         Top             =   570
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "About"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   7440
         Max             =   100
         TabIndex        =   14
         Top             =   6420
         Value           =   75
         Width           =   1695
      End
      Begin Project1.XPButton XPButton5 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   6360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Play"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   6000
         Width           =   8895
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   3150
         Left            =   2520
         TabIndex        =   5
         ToolTipText     =   "Double click on one of the songs to play them."
         Top             =   2760
         Width           =   6615
      End
      Begin Project1.XPButton XPButton4 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   5280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mix"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.XPButton XPButton3 
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   4680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Urankaatha Vizhikal"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.XPButton XPButton2 
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   4080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Deva"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.XPButton XPButton1 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   3480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "A.R.Rahman"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.XPButton XPButton7 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   6360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Pause"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.XPButton XPButton6 
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   6360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Stop"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on a catergories below and select a song from the list to play "
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   9120
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   6840
         Picture         =   "Form1.frx":B5F6
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   6840
         Picture         =   "Form1.frx":C2C0
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   4200
         Picture         =   "Form1.frx":CF8A
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   4200
         Picture         =   "Form1.frx":DC54
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1440
         Picture         =   "Form1.frx":E91E
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1440
         Picture         =   "Form1.frx":F5E8
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "My Explorer"
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
         Left            =   7440
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "My Tamil Keyboard"
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
         Left            =   7440
         TabIndex        =   21
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "My Web Browser"
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
         Left            =   4800
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "My Media Player"
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
         Left            =   4800
         TabIndex        =   19
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "My Text Editor"
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
         Left            =   2040
         TabIndex        =   18
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "My Screen Capture"
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
         Left            =   2040
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dhayalan Sivasuthan's some other products:"
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "Form1.frx":102B2
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   960
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   1320
         X2              =   1320
         Y1              =   6480
         Y2              =   6720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   6480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   6480
         Visible         =   0   'False
         Width           =   975
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   720
         Left            =   360
         TabIndex        =   11
         Top             =   5160
         Visible         =   0   'False
         Width           =   2220
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
         _cx             =   3916
         _cy             =   1270
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   8535
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tamil Music Collection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub File1_Click()
WindowsMediaPlayer1.URL = File1.Path + "\" + File1.FileName
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim WindowRegion As Long
    
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Set picMainSkin.Picture = LoadPicture(App.Path & "\tuneup skin 2.gif")
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.Hwnd, WindowRegion, True

File1.Path = App.Path + "\" + "Rahman\"
WindowsMediaPlayer1.URL = App.Path + "\" + "Mix\Billa Theme.mp3"
End Sub



Private Sub Form_Unload(Cancel As Integer)
Dialog.Show
Me.Enabled = False
End Sub

Private Sub HScroll1_Scroll()
Timer2.Enabled = False
WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value
Timer2.Enabled = True

End Sub



Private Sub HScroll2_Change()
WindowsMediaPlayer1.settings.volume = HScroll2.Value

End Sub

Private Sub HScroll2_Scroll()
WindowsMediaPlayer1.settings.volume = HScroll2.Value

End Sub

Private Sub Image1_Click()
DEx = True
Dialog.Show
Me.Enabled = False
End Sub

Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.Hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub



Private Sub Timer1_Timer()
If WindowsMediaPlayer1.URL = "" Then Exit Sub
Label2.Caption = WindowsMediaPlayer1.Controls.currentPositionString
'& "/" &
Label3.Caption = WindowsMediaPlayer1.currentMedia.durationString
'If WindowsMediaPlayer1.currentMedia.duration < 5 Then Exit Sub
'If WindowsMediaPlayer1.Controls.currentPosition > WindowsMediaPlayer1.currentMedia.duration - 5 Then
'imgNext_Click
'End If

End Sub

Private Sub Timer2_Timer()
HScroll1.Value = WindowsMediaPlayer1.Controls.currentPosition
End Sub

Private Sub WindowsMediaPlayer1_MediaChange(ByVal Item As Object)
HScroll1.Max = WindowsMediaPlayer1.currentMedia.duration
End Sub




Private Sub XPButton1_Click()
File1.Path = App.Path + "\" + "Rahman\"
End Sub

Private Sub XPButton2_Click()
File1.Path = App.Path + "\" + "Deva\"

End Sub

Private Sub XPButton3_Click()
File1.Path = App.Path + "\" + "Urankaatha Vizhikal\"

End Sub

Private Sub XPButton4_Click()
File1.Path = App.Path + "\" + "Mix\"

End Sub

Private Sub XPButton5_Click()
WindowsMediaPlayer1.Controls.play
End Sub

Private Sub XPButton6_Click()
WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub XPButton7_Click()
WindowsMediaPlayer1.Controls.pause
End Sub

Private Sub XPButton8_Click()
Me.Enabled = False
Dialog.Show
End Sub
