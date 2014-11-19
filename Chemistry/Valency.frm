VERSION 5.00
Begin VB.Form frmValency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valency"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   Icon            =   "Valency.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Valency.frx":57E2
   ScaleHeight     =   4845
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Tag             =   "s"
   Begin VB.TextBox K 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   36
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox K2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Si 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   53
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Si2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Hg2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox Cu2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox Zn2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Xe2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox S2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox Na2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Ag2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox P2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox Sn2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox O2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox N2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   7920
      Width           =   495
   End
   Begin VB.TextBox Mg2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Li2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Pb2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox Kr2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Fe2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox I2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox H2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Au2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox F2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Cl2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox C2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Ca2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Br2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Ba2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Al2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Results"
      Default         =   -1  'True
      DownPicture     =   "Valency.frx":C757
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4200
      Width           =   5655
   End
   Begin VB.TextBox Zn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   52
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Xe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   51
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text24 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox S 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox Na 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   48
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Ag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   47
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox P 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox O 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   45
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox N 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   7920
      Width           =   495
   End
   Begin VB.TextBox Hg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox Mg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   42
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Li 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   41
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Pb 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox Kr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   39
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Fe 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox I 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   37
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox H 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   35
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Au 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   34
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox F 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   33
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Cu 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox Cl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   31
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox C 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   30
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Ca 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   29
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Br 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   28
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Ba 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   27
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Al 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   26
      Top             =   600
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type the valency of the element in the textboxes in front of them"
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
      Left            =   120
      TabIndex        =   113
      Top             =   120
      Width           =   5655
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   480
      Y2              =   4080
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Potassium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   112
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   27
      Left            =   2520
      TabIndex        =   111
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Silicon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   109
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   26
      Left            =   5640
      TabIndex        =   108
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   25
      Left            =   5640
      TabIndex        =   80
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   24
      Left            =   5640
      TabIndex        =   79
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   23
      Left            =   2760
      TabIndex        =   78
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   22
      Left            =   2760
      TabIndex        =   77
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   21
      Left            =   5640
      TabIndex        =   76
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   20
      Left            =   5640
      TabIndex        =   75
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   19
      Left            =   2760
      TabIndex        =   74
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   18
      Left            =   5640
      TabIndex        =   73
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   2760
      TabIndex        =   72
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   16
      Left            =   2760
      TabIndex        =   71
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   15
      Left            =   5640
      TabIndex        =   70
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   14
      Left            =   5640
      TabIndex        =   69
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   13
      Left            =   2760
      TabIndex        =   68
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   5640
      TabIndex        =   67
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   66
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   65
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   64
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   63
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   62
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "AA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   61
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   60
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   59
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   58
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   57
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   56
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label AA 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   55
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Chlorine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Copper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Fluorine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Gold"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Hydrogen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Iodine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Iron"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Krypton"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "Zinc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "Xenon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Tin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Sulphur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Sodium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Silver"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Phosphorus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Oxygen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Carbon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Calcium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bromine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Barium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aluminium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Lead"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Lithium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Magnesium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Nitrogen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Mercury"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "frmValency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Al2.Visible = True
Ba2.Visible = True
Br2.Visible = True
Ca2.Visible = True
C2.Visible = True
Cl2.Visible = True
F2.Visible = True
Au2.Visible = True
H2.Visible = True
K2.Visible = True
I2.Visible = True
Kr2.Visible = True
Li2.Visible = True
Mg2.Visible = True
O2.Visible = True
Ag2.Visible = True
Na2.Visible = True
Xe2.Visible = True
Zn2.Visible = True
Si2.Visible = True



AA(0).Visible = True
AA(1).Visible = True
AA(2).Visible = True
AA(3).Visible = True
AA(4).Visible = True
AA(5).Visible = True
AA(6).Visible = True
AA(7).Visible = True
AA(8).Visible = True
AA(9).Visible = True
AA(10).Visible = True
AA(11).Visible = True
AA(12).Visible = True
AA(13).Visible = True
AA(14).Visible = True
AA(15).Visible = True
AA(16).Visible = True
AA(17).Visible = True
AA(18).Visible = True
AA(19).Visible = True
AA(20).Visible = True
AA(21).Visible = True
AA(22).Visible = True
AA(23).Visible = True
AA(24).Visible = True
AA(25).Visible = True
AA(26).Visible = True
AA(27).Visible = True



























If Al = "3" Then
Al2 = "C"
Else
Al2 = "W"
End If

If Ba = "2" Then
Ba2 = "C"
Else
Ba2 = "W"
End If

If Br = "2" Then
Br2 = "C"
Else
Br2 = "W"
End If

If Ca = "2" Then
Ca2 = "C"
Else
Ca2 = "W"
End If

If Cl = "1" Then
Cl2 = "C"
Else
Cl2 = "W"
End If

If C = "4" Then
C2 = "C"
Else
C2 = "W"
End If

If F = "1" Then
F2 = "C"
Else
F2 = "W"
End If

If H = "1" Then
H2 = "C"
Else
H2 = "W"
End If

If K = "1" Then
K2 = "C"
Else
K2 = "W"
End If

If I = "1" Then
I2 = "C"
Else
I2 = "W"
End If

If Kr = "0" Then
Kr2 = "C"
Else
Kr2 = "W"
End If

If Li = "1" Then
Li2 = "C"
Else
Li2 = "W"
End If

If Mg = "2" Then
Mg2 = "C"
Else
Mg2 = "W"
End If

If O = "2" Then
O2 = "C"
Else
O2 = "W"
End If


If Ag = "1" Then
Ag2 = "C"
Else
Ag2 = "W"
End If

If Na = "1" Then
Na2 = "C"
Else
Na2 = "W"
End If


If Xe = "0" Then
Xe2 = "C"
Else
Xe2 = "W"
End If

If Zn = "2" Then
Zn2 = "C"
Else
Zn2 = "W"
End If

If Si = "2" Then
Si2 = "C"
Else
Si2 = "W"
End If
End Sub

Private Sub Command2_Click()
Unload Me
 

MsgBox "Once again, Good Luck", vbOKOnly, "Siva"
frmValency.Show
End Sub

Private Sub Form_Load()
MsgBox "Good Luck", vbOKOnly, "Siva"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide

frmWelcomescreen.Show
End Sub
