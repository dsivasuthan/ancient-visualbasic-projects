VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tamil Keyboard - Dhayalan Sivasuthan"
   ClientHeight    =   6345
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   7800
   DrawMode        =   7  'Invert
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Wingdings"
      Size            =   10.5
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "TamilKeyboardForm1.frx":0000
   LinkTopic       =   "Translator"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6345
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   960
      TabIndex        =   81
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7560
      Top             =   3960
   End
   Begin VB.CommandButton Command15 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5760
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TamilKeyboardForm1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Tip"
      Top             =   4800
      WhatsThisHelpID =   2
      Width           =   1935
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H8000000D&
      Caption         =   "nfhgp"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Copy Text to clipboard"
      Top             =   120
      WhatsThisHelpID =   2
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "vd;lu;"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "Erase:whole textbox"
      Top             =   4560
      WhatsThisHelpID =   2
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "kP"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   77
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdMana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   4920
      MaskColor       =   &H008080FF&
      TabIndex        =   7
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H000040C0&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5040
      TabIndex        =   47
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdDeeyanna 
      BackColor       =   &H000040C0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5040
      TabIndex        =   46
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   120
      MaskColor       =   &H00404040&
      TabIndex        =   21
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton Command10 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4560
      TabIndex        =   74
      Top             =   5205
      Width           =   612
   End
   Begin VB.CommandButton Command9 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   73
      Top             =   5205
      Width           =   612
   End
   Begin VB.CommandButton Command8 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      TabIndex        =   72
      Top             =   5205
      Width           =   612
   End
   Begin VB.CommandButton Command7 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   71
      Top             =   5205
      Width           =   612
   End
   Begin VB.CommandButton Command6 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   70
      Top             =   5205
      Width           =   612
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H80000010&
      Cancel          =   -1  'True
      Caption         =   ",ilntsp tpL"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "Space"
      ToolTipText     =   "Space"
      Top             =   5760
      Width           =   5532
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0C0&
      Caption         =   "mop"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Erase:whole textbox"
      Top             =   1800
      WhatsThisHelpID =   2
      Width           =   1212
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00E0E0E0&
      Caption         =   "%L"
      DisabledPicture =   "TamilKeyboardForm1.frx":2994
      DragIcon        =   "TamilKeyboardForm1.frx":F197
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      MaskColor       =   &H000000FF&
      MouseIcon       =   "TamilKeyboardForm1.frx":F5D9
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Close"
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      HelpContextID   =   2
      Left            =   120
      MouseIcon       =   "TamilKeyboardForm1.frx":14DBB
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   75
      Top             =   120
      WhatsThisHelpID =   2
      Width           =   6252
   End
   Begin VB.CommandButton cmdNana3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   10
      Left            =   4320
      MaskColor       =   &H008080FF&
      TabIndex        =   0
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdRara 
      BackColor       =   &H00FFC0FF&
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   3720
      MaskColor       =   &H008080FF&
      TabIndex        =   67
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdLana3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   3120
      MaskColor       =   &H008080FF&
      TabIndex        =   2
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdLana2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   8
      Left            =   2520
      MaskColor       =   &H008080FF&
      TabIndex        =   3
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdVana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   1920
      MaskColor       =   &H008080FF&
      TabIndex        =   5
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdLana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   1320
      MaskColor       =   &H008080FF&
      TabIndex        =   4
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdRana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   6
      Left            =   720
      MaskColor       =   &H008080FF&
      TabIndex        =   6
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdYana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   120
      MaskColor       =   &H008080FF&
      TabIndex        =   8
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdPana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   4920
      MaskColor       =   &H008080FF&
      TabIndex        =   9
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdNana2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   4320
      MaskColor       =   &H008080FF&
      TabIndex        =   11
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdThana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "j"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   3720
      MaskColor       =   &H008080FF&
      TabIndex        =   10
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdNana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   3120
      MaskColor       =   &H008080FF&
      TabIndex        =   12
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdDana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   2520
      MaskColor       =   &H008080FF&
      TabIndex        =   13
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdGnana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   1920
      MaskColor       =   &H008080FF&
      TabIndex        =   14
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdSana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   1320
      MaskColor       =   &H008080FF&
      TabIndex        =   17
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdGana 
      BackColor       =   &H00FFC0FF&
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   720
      MaskColor       =   &H008080FF&
      TabIndex        =   16
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdKanaa 
      BackColor       =   &H00FFC0FF&
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      MaskColor       =   &H008080FF&
      TabIndex        =   15
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdNoona3 
      BackColor       =   &H000080FF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   11
      Left            =   4320
      TabIndex        =   30
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdLoona3 
      BackColor       =   &H000080FF&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   3720
      TabIndex        =   32
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdLoona2 
      BackColor       =   &H000080FF&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   9
      Left            =   3120
      TabIndex        =   33
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdVoona 
      BackColor       =   &H000080FF&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   2520
      TabIndex        =   34
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdRoona 
      BackColor       =   &H000080FF&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   1920
      TabIndex        =   36
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdLoona 
      BackColor       =   &H000080FF&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   1320
      TabIndex        =   35
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdYoona 
      BackColor       =   &H000080FF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   720
      TabIndex        =   37
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdMoona 
      BackColor       =   &H000080FF&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   120
      TabIndex        =   39
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton cmdTroona 
      BackColor       =   &H000080FF&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   4320
      TabIndex        =   31
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton cmdPoona 
      BackColor       =   &H000080FF&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   3720
      TabIndex        =   40
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton cmdNoona2 
      BackColor       =   &H000080FF&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   3120
      TabIndex        =   41
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton cmdThoona 
      BackColor       =   &H000080FF&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   2520
      TabIndex        =   42
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton cmdNoona 
      BackColor       =   &H000080FF&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   1920
      TabIndex        =   43
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton cmdDoona 
      BackColor       =   &H000080FF&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   1320
      TabIndex        =   45
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Soona 
      BackColor       =   &H000080FF&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   720
      TabIndex        =   49
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton cmdKoona 
      BackColor       =   &H000080FF&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   3720
      Width           =   612
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H0080C0FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   2760
      TabIndex        =   38
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080C0FF&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   2160
      TabIndex        =   44
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   1560
      TabIndex        =   48
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton cmdKoovanna 
      BackColor       =   &H0080C0FF&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   960
      TabIndex        =   52
      Top             =   4560
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6840
      TabIndex        =   54
      Top             =   4200
      Width           =   492
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H00FF8080&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6360
      TabIndex        =   63
      Top             =   4200
      Width           =   492
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FF8080&
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5880
      TabIndex        =   55
      Top             =   4200
      Width           =   492
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H00FF8080&
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6840
      TabIndex        =   60
      Top             =   3840
      Width           =   492
   End
   Begin VB.CommandButton Command47 
      BackColor       =   &H00FF8080&
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6360
      TabIndex        =   62
      Top             =   3840
      Width           =   492
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H00FF8080&
      Caption         =   "h;"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5880
      TabIndex        =   61
      Top             =   3840
      Width           =   492
   End
   Begin VB.CommandButton Command50 
      BackColor       =   &H00FF8080&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6840
      TabIndex        =   59
      Top             =   3480
      Width           =   492
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00FF8080&
      Caption         =   "hP"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6360
      TabIndex        =   65
      Top             =   3480
      Width           =   492
   End
   Begin VB.CommandButton Command51 
      BackColor       =   &H00FF8080&
      Caption         =   "up"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5880
      TabIndex        =   58
      ToolTipText     =   "Visiri"
      Top             =   3480
      Width           =   492
   End
   Begin VB.CommandButton Command52 
      BackColor       =   &H00FF8080&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   6360
      TabIndex        =   57
      Top             =   3120
      Width           =   492
   End
   Begin VB.CommandButton h 
      BackColor       =   &H00FF8080&
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5880
      TabIndex        =   64
      ToolTipText     =   "Aravu"
      Top             =   3120
      Width           =   492
   End
   Begin VB.CommandButton cmdOhvanna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Xs"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   6960
      MaskColor       =   &H00404040&
      TabIndex        =   66
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   732
   End
   Begin VB.CommandButton cmdOhvanna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   6240
      MaskColor       =   &H00404040&
      TabIndex        =   29
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   732
   End
   Begin VB.CommandButton Oahna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5520
      MaskColor       =   &H00404040&
      TabIndex        =   28
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   732
   End
   Begin VB.CommandButton Iyyanna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4920
      MaskColor       =   &H00404040&
      TabIndex        =   22
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdAeanna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      MaskColor       =   &H00404040&
      TabIndex        =   23
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdAena 
      BackColor       =   &H00C0C0C0&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      MaskColor       =   &H00404040&
      TabIndex        =   27
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdOovanna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      MaskColor       =   &H00404040&
      TabIndex        =   18
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdOona 
      BackColor       =   &H00C0C0C0&
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      MaskColor       =   &H00404040&
      TabIndex        =   26
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdEyanna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      MaskColor       =   &H00404040&
      TabIndex        =   20
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdEena 
      BackColor       =   &H00C0C0C0&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      MaskColor       =   &H00404040&
      TabIndex        =   19
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdAavanna 
      BackColor       =   &H00C0C0C0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   720
      MaskColor       =   &H00404040&
      TabIndex        =   25
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton cmdAnaa 
      BackColor       =   &H00C0C0C0&
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      MaskColor       =   &H00404040&
      TabIndex        =   24
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   612
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   960
      TabIndex        =   69
      Top             =   5205
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   68
      Top             =   5205
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   120
      TabIndex        =   76
      Top             =   4900
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      DragIcon        =   "TamilKeyboardForm1.frx":16405
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Alankaram"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   50
      Top             =   2745
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub cmdAavanna_Click()
Text1.Text = Text1.Text + "M"
End Sub

Private Sub cmdAeanna_Click()
Text1.Text = Text1.Text + "V"
End Sub

Private Sub cmdAena_Click()
Text1.Text = Text1.Text + "v"
End Sub

Private Sub cmdAnaa_Click()
Text1.Text = Text1.Text + "m"
End Sub

Private Sub cmdDana_Click(Index As Integer)
Text1.Text = Text1.Text + "l"
End Sub

Private Sub cmdDeeyanna_Click(Index As Integer)
Text1.Text = Text1.Text + "B"
End Sub

Private Sub cmdDoona_Click(Index As Integer)
Text1.Text = Text1.Text + "L"
End Sub

Private Sub cmdEena_Click()
Text1.Text = Text1.Text + ","
End Sub

Private Sub cmdEyanna_Click()
Text1.Text = Text1.Text + "<"
End Sub

Private Sub cmdGana_Click(Index As Integer)
Text1.Text = Text1.Text + "q"
End Sub

Private Sub cmdGnana_Click(Index As Integer)
Text1.Text = Text1.Text + "Q"
End Sub

Private Sub cmdKanaa_Click(Index As Integer)
Text1.Text = Text1.Text + "f"
End Sub

Private Sub cmdKoona_Click(Index As Integer)
Text1.Text = Text1.Text + "F"
End Sub

Private Sub cmdKoovanna_Click(Index As Integer)
Text1.Text = Text1.Text + "$"
End Sub

Private Sub cmdLana_Click(Index As Integer)
Text1.Text = Text1.Text + "y"
End Sub

Private Sub cmdLana2_Click(Index As Integer)
Text1.Text = Text1.Text + "o"
End Sub

Private Sub cmdLana3_Click(Index As Integer)
Text1.Text = Text1.Text + "s"
End Sub

Private Sub cmdLoona_Click(Index As Integer)
Text1.Text = Text1.Text + "Y"
End Sub

Private Sub cmdLoona2_Click(Index As Integer)
Text1.Text = Text1.Text + "O"
End Sub

Private Sub cmdLoona3_Click(Index As Integer)
Text1.Text = Text1.Text + "S"
End Sub

Private Sub cmdMana_Click(Index As Integer)
Text1.Text = Text1.Text + "k"
End Sub

Private Sub cmdMoona_Click(Index As Integer)
Text1.Text = Text1.Text + "K"
End Sub

Private Sub cmdNana_Click(Index As Integer)
Text1.Text = Text1.Text + "z"
End Sub

Private Sub cmdNana2_Click(Index As Integer)
Text1.Text = Text1.Text + "e"
End Sub

Private Sub cmdNana3_Click(Index As Integer)
Text1.Text = Text1.Text + "d"
End Sub

Private Sub cmdNoona_Click(Index As Integer)
Text1.Text = Text1.Text + "Z"
End Sub

Private Sub cmdNoona2_Click(Index As Integer)
Text1.Text = Text1.Text + "E"
End Sub

Private Sub cmdNoona3_Click(Index As Integer)
Text1.Text = Text1.Text + "D"
End Sub

Private Sub cmdOhvanna_Click(Index As Integer)
Text1.Text = Text1.Text + "Xs"
End Sub

Private Sub cmdOona_Click()
Text1.Text = Text1.Text + "c"
End Sub

Private Sub cmdOovanna_Click()
Text1.Text = Text1.Text + "C"
End Sub

Private Sub cmdPana_Click(Index As Integer)
Text1.Text = Text1.Text + "g"
End Sub

Private Sub cmdPoona_Click(Index As Integer)
Text1.Text = Text1.Text + "G"
End Sub

Private Sub cmdRana_Click(Index As Integer)
Text1.Text = Text1.Text + "u"
End Sub

Private Sub cmdRara_Click(Index As Integer)
Text1.Text = Text1.Text + "w"
End Sub

Private Sub cmdRoona_Click(Index As Integer)
Text1.Text = Text1.Text + "U"
End Sub

Private Sub cmdSana_Click(Index As Integer)
Text1.Text = Text1.Text + "r"
End Sub

Private Sub cmdThana_Click(Index As Integer)
Text1.Text = Text1.Text + "j"
End Sub

Private Sub cmdThoona_Click(Index As Integer)
Text1.Text = Text1.Text + "J"
End Sub

Private Sub cmdTroona_Click(Index As Integer)
Text1.Text = Text1.Text + "W"
End Sub

Private Sub cmdVana_Click(Index As Integer)
Text1.Text = Text1.Text + "t"
End Sub

Private Sub cmdVoona_Click(Index As Integer)
Text1.Text = Text1.Text + "T"
End Sub

Private Sub cmdYana_Click(Index As Integer)
Text1.Text = Text1.Text + "a"
End Sub

Private Sub cmdYoona_Click(Index As Integer)
Text1.Text = Text1.Text + "A"
End Sub

Private Sub Command1_Click()
Text1.Text = Text1.Text + "n"
End Sub

Private Sub Command10_Click()
Text1.Text = Text1.Text + "%-"
End Sub

Private Sub Command11_Click(Index As Integer)
Text1.Text = Text1.Text + "^"
End Sub

Private Sub Command12_Click()
Text1.Text = Text1.Text + "P"
End Sub

Private Sub Command13_Click()
Clipboard.SetText Text1.Text
End Sub

Private Sub Command14_Click(Index As Integer)
Text1.Text = Text1.Text + "b"
End Sub

Private Sub Command15_Click()
Dialog.Show
Me.Enabled = False

End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text + vbNewLine

End Sub

Private Sub Command20_Click()
Text1.Text = Text1.Text + " "
End Sub



Private Sub Command20_KeyPress(KeyAscii As Integer)
Text1.Text = Text1.Text + " "
End Sub

Private Sub Command21_Click(Index As Integer)
Text1.Text = Text1.Text + "i"
End Sub

Private Sub Command23_Click(Index As Integer)
Text1.Text = Text1.Text + "%"
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text + "."
End Sub

Private Sub Command31_Click()
intRes = MsgBox("This will erase all the text in the textbox.! Do you realy want to continue!?", vbYesNo, "Clear Textbox!!??")
If intRes = vbYes Then
Text1 = ""
End If
End Sub

Private Sub Command32_Click()
intRes = MsgBox("Do you really want to close The Translator?!", vbYesNo, "Are You Sure?!")
If intRes = vbYes Then
Dialog.Show
Timer1.Enabled = True
End If
End Sub

Private Sub Command36_Click()
Text1.Text = Text1.Text + "/"
End Sub

Private Sub Command4_Click(Index As Integer)
Text1.Text = Text1.Text + "#"
End Sub

Private Sub Command44_Click(Index As Integer)
Text1.Text = Text1.Text + "P"
End Sub

Private Sub Command46_Click(Index As Integer)
Text1.Text = Text1.Text + "N"
End Sub

Private Sub Command47_Click(Index As Integer)
Text1.Text = Text1.Text + "}"
End Sub

Private Sub Command48_Click(Index As Integer)
Text1.Text = Text1.Text + ";"
End Sub

Private Sub Command49_Click(Index As Integer)
Text1.Text = Text1.Text + "{"
End Sub


Private Sub Command5_Click()
Text1.Text = Text1.Text + ">"
End Sub

Private Sub Command50_Click(Index As Integer)
Text1.Text = Text1.Text + "_"
End Sub

Private Sub Command51_Click(Index As Integer)
Text1.Text = Text1.Text + "p"
End Sub

Private Sub Command52_Click(Index As Integer)
Text1.Text = Text1.Text + "+"
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text + "?"
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text + "!"
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text + "("
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + ")"
End Sub

Private Sub Form_Load()
MsgBox "Welcome to Siva's Tamil Keyboard!  To use this program you must have the tamil font Alankaram", vbOKOnly, "Siva"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Command32_Click
End Sub

Private Sub h_Click(Index As Integer)
Text1.Text = Text1.Text + "h"
End Sub

Private Sub Iyyanna_Click()
Text1.Text = Text1.Text + "I"
End Sub

Private Sub Oahna_Click()
Text1.Text = Text1.Text + "x"
End Sub

Private Sub Soona_Click(Index As Integer)
Text1.Text = Text1.Text + "R"
End Sub

Private Sub Timer1_Timer()
If Me.Height > 600 Then
Me.Height = Me.Height - 50
Else
End
End If
End Sub
