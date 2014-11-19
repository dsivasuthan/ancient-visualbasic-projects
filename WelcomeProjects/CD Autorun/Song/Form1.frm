VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
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
   ScaleHeight     =   7410
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2160
      ScaleHeight     =   1515
      ScaleWidth      =   2955
      TabIndex        =   19
      Top             =   7440
      Width           =   3015
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Dhayalan Sivasuthan's some other products:"
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label17 
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
         Left            =   1800
         TabIndex        =   25
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label16 
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
         Left            =   1800
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label15 
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
         Left            =   4560
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label14 
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
         Left            =   4560
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label13 
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
         Left            =   7200
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label12 
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
         Left            =   7200
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.Image Image14 
         Height          =   480
         Left            =   1200
         Picture         =   "Form1.frx":57E2
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   1200
         Picture         =   "Form1.frx":64AC
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   3960
         Picture         =   "Form1.frx":7176
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   3960
         Picture         =   "Form1.frx":7E40
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   6600
         Picture         =   "Form1.frx":8B0A
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   6600
         Picture         =   "Form1.frx":97D4
         Top             =   840
         Width           =   480
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   8880
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6480
      Top             =   5040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6960
      Top             =   4440
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
      Picture         =   "Form1.frx":A49E
      ScaleHeight     =   7035
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   2520
         TabIndex        =   27
         ToolTipText     =   "Double click on one of the songs to play them."
         Top             =   2280
         Width           =   6615
      End
      Begin Project1.XPButton XPButton8 
         Height          =   300
         Left            =   5400
         TabIndex        =   10
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
         ForeHover       =   32768
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   7440
         Max             =   100
         TabIndex        =   9
         Top             =   6420
         Value           =   75
         Width           =   1695
      End
      Begin Project1.XPButton XPButton5 
         Height          =   375
         Left            =   240
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   6000
         Width           =   8895
      End
      Begin Project1.XPButton XPButton7 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
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
         TabIndex        =   5
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
      Begin Project1.XPButton XPButton4 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   3360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mix 1"
         ForeColor       =   -2147483642
         ForeHover       =   65280
      End
      Begin Project1.XPButton XPButton3 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
         ForeHover       =   12582912
      End
      Begin Project1.XPButton XPButton2 
         Height          =   315
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
         ForeHover       =   16576
      End
      Begin Project1.XPButton XPButton1 
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
         ForeHover       =   49152
      End
      Begin Project1.XPButton XPButton9 
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Top             =   3720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mix 2"
         ForeColor       =   -2147483642
         ForeHover       =   16711680
      End
      Begin Project1.XPButton XPButton10 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   4080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mix 3"
         ForeColor       =   -2147483642
         ForeHover       =   192
      End
      Begin Project1.XPButton XPButton11 
         Height          =   315
         Left            =   240
         TabIndex        =   34
         Top             =   4440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mix 4"
         ForeColor       =   -2147483642
         ForeHover       =   8438015
      End
      Begin Project1.XPButton XPButton12 
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Top             =   4800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mix 5"
         ForeColor       =   -2147483642
         ForeHover       =   16576
      End
      Begin Project1.XPButton XPButton13 
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Top             =   5160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Vijay"
         ForeColor       =   -2147483642
         ForeHover       =   12632256
      End
      Begin Project1.XPButton XPButton14 
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   5520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
         ForeHover       =   16384
      End
      Begin VB.Image Image15 
         Height          =   540
         Left            =   4440
         Picture         =   "Form1.frx":102B2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on a catergories button below and select a song from the list to play!"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   1920
         Width           =   5775
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   960
         Picture         =   "Form1.frx":22042
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   720
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
         TabIndex        =   8
         Top             =   6480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   6480
         Visible         =   0   'False
         Width           =   975
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   15
         Left            =   360
         TabIndex        =   6
         Top             =   5160
         Visible         =   0   'False
         Width           =   15
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
         uiMode          =   "invisible"
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
         _cx             =   26
         _cy             =   26
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
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan's some other products:"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   0
      Width           =   6015
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
      Left            =   1800
      TabIndex        =   17
      Top             =   480
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
      Left            =   1800
      TabIndex        =   16
      Top             =   960
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
      Left            =   4560
      TabIndex        =   15
      Top             =   480
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
      Left            =   4560
      TabIndex        =   14
      Top             =   960
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
      Left            =   7200
      TabIndex        =   13
      Top             =   480
      Width           =   1815
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
      Left            =   7200
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   1200
      Picture         =   "Form1.frx":23B84
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1200
      Picture         =   "Form1.frx":2484E
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   3960
      Picture         =   "Form1.frx":25518
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   3960
      Picture         =   "Form1.frx":261E2
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   6600
      Picture         =   "Form1.frx":26EAC
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   6600
      Picture         =   "Form1.frx":27B76
      Top             =   840
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8880
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub File1_Click()
On Error Resume Next
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
If Right(App.Path, 1) <> "\" Then
File1.Path = App.Path + "\" + "Rahman\"
Else
File1.Path = App.Path + "Rahman\"
End If

If Right(App.Path, 1) <> "\" Then
WindowsMediaPlayer1.URL = App.Path + "\" + "Mix 1\Billa Theme.mp3"
Else
WindowsMediaPlayer1.URL = App.Path + "Mix 1\Billa Theme.mp3"
End If
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
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Rahman\"
Else
File1.Path = App.Path + "\" + "Rahman\"
End If


End Sub

Private Sub XPButton10_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Mix 3\"
Else
File1.Path = App.Path + "\" + "Mix 3\"
End If
End Sub

Private Sub XPButton11_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Mix 4\"
Else
File1.Path = App.Path + "\" + "Mix 4\"
End If
End Sub

Private Sub XPButton12_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Mix 5\"
Else
File1.Path = App.Path + "\" + "Mix 5\"
End If
End Sub

Private Sub XPButton13_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Vijay\"
Else
File1.Path = App.Path + "\" + "Vijay\"
End If
End Sub

Private Sub XPButton14_Click()
Me.Enabled = False
Dialog.Show

End Sub

Private Sub XPButton2_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Deva\"
Else
File1.Path = App.Path + "\" + "Deva\"
End If

End Sub

Private Sub XPButton3_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Urankaatha Vizhikal\"
Else
File1.Path = App.Path + "\" + "Urankaatha Vizhikal\"
End If

End Sub

Private Sub XPButton4_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Mix 1\"
Else
File1.Path = App.Path + "\" + "Mix 1\"
End If

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

Private Sub XPButton9_Click()
If Right(App.Path, 1) = "\" Then
File1.Path = App.Path + "Mix 2\"
Else
File1.Path = App.Path + "\" + "Mix 2\"
End If
End Sub
